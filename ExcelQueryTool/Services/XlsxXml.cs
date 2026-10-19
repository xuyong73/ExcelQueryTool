using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace ExcelQueryTool.Services;

/// <summary>
/// xlsx 包内部件的 XML 读取入口。
///
/// 为什么必须显式固定 UTF-8 编码、不能让 XmlReader 自行嗅探：
/// xlsx 包内部件的 XML 声明恒为 encoding="UTF-8"，但 ZipArchiveEntry.Open() 返回的是
/// 无缓冲的顺序解压流。XmlReader 的编码嗅探依赖流开头的若干字节，在这种流上会因
/// 前导字节不足而误判 —— 结果是 UTF-8 字节被按错误编码组字符，中文变成乱码
/// （实测表现为 t="s" 的共享字符串单元格整体失效）。
///
/// 这里统一绕开嗅探：如果部件自身带 BOM 则按 BOM 处理，否则一律按 UTF-8 解码；
/// 同时给大部件套上缓冲，避免逐字节读取解压流。
/// </summary>
internal static class XlsxXml
{
    /// <summary>常见部件的缓冲尺寸（命名常量，消除散落的 1&lt;&lt;14 / 1&lt;&lt;16 / 1&lt;&lt;20 魔法数）</summary>
    public const int BufferSmall = 1 << 14;    // styles / workbook / rels 等小部件
    public const int BufferMedium = 1 << 16;   // sharedStrings / col
    public const int BufferSheet = 1 << 20;    // sheet 部件可能数百 MB

    /// <summary>
    /// 打开 zip 内部件 <paramref name="partPath"/> 并创建部件读取器；部件不存在返回 null。
    /// 读取器关闭时连带关闭底层流（<c>CloseInput = true</c>）。
    /// </summary>
    public static XmlReader? OpenPartReader(this ZipArchive zip, string partPath, int bufferSize)
    {
        var entry = zip.GetEntry(partPath);
        if (entry == null) return null;
        return CreatePartReader(entry.Open(), bufferSize);
    }

    public static XmlReader CreatePartReader(Stream partStream, int bufferSize = 0)
    {
        ArgumentNullException.ThrowIfNull(partStream);

        var stream = partStream;

        // 带 BOM 时交给 StreamReader 依 BOM 判定，否则强制 UTF-8
        var encoding = DetectBom(stream) ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        var textReader = new StreamReader(
            bufferSize > 0 ? new BufferedStream(stream, bufferSize) : stream,
            encoding,
            detectEncodingFromByteOrderMarks: true);

        return XmlReader.Create(textReader, new XmlReaderSettings
        {
            IgnoreWhitespace = true,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            DtdProcessing = DtdProcessing.Prohibit,
            CloseInput = true,
        });
    }

    /// <summary>探测 BOM；无 BOM 返回 null</summary>
    private static Encoding? DetectBom(Stream stream)
    {
        if (!stream.CanSeek) return null;

        long origin = stream.Position;
        Span<byte> head = stackalloc byte[4];
        int read = stream.Read(head);
        stream.Position = origin;

        if (read >= 3 && head[0] == 0xEF && head[1] == 0xBB && head[2] == 0xBF) return new UTF8Encoding(false);
        if (read >= 4 && head[0] == 0xFF && head[1] == 0xFE && head[2] == 0x00 && head[3] == 0x00) return new UTF32Encoding(false, true);
        if (read >= 4 && head[0] == 0x00 && head[1] == 0x00 && head[2] == 0xFE && head[3] == 0xFF) return new UTF32Encoding(true, true);
        if (read >= 2 && head[0] == 0xFF && head[1] == 0xFE) return Encoding.Unicode;
        if (read >= 2 && head[0] == 0xFE && head[1] == 0xFF) return Encoding.BigEndianUnicode;

        return null;
    }
}

/// <summary>
/// OOXML 关系（.rels）的共享实现：命名空间/关系类型常量 + 流式 .rels 解析 + workbook sheet 枚举。
///
/// <see cref="XlsxPictureIndex"/> 与 <see cref="XlsxPackageLocator"/> 历史上各持一份
/// NsDocProps/RelType* 常量与 .rels 解析（DOM vs 流式两种风格），靠注释同步；
/// 本类是唯一一份实现（与 <see cref="XlsxPartPath"/> 收敛包路径是同一思路）。
/// </summary>
internal static class XlsxRelationships
{
    /// <summary>关系命名空间（r:id / r:embed / r:link 等属性的前缀命名空间）。</summary>
    public const string NsRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>workbook rels 中工作表关系的 Type。</summary>
    public const string RelTypeWorksheet = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

    /// <summary>sheet rels 中 drawing 部件关系的 Type。</summary>
    public const string RelTypeDrawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

    /// <summary>richData / drawing rels 中图片关系的 Type。</summary>
    public const string RelTypeImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    /// <summary>
    /// 读取部件的 .rels，返回 relId → (Type, Target)。
    /// 流式解析（与 <see cref="XlsxXml.CreatePartReader"/> 同款缓冲），部件不存在返回空字典。
    /// </summary>
    public static Dictionary<string, (string Type, string Target)> ReadRels(ZipArchive zip, string partPath)
    {
        var result = new Dictionary<string, (string Type, string Target)>(StringComparer.Ordinal);
        var entry = zip.GetEntry(XlsxPartPath.BuildRelsPath(partPath));
        if (entry == null) return result;

        using var stream = entry.Open();
        using var reader = XlsxXml.CreatePartReader(stream, XlsxXml.BufferSmall);

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship") continue;

            string? id = reader.GetAttribute("Id");
            string? type = reader.GetAttribute("Type");
            string? target = reader.GetAttribute("Target");
            if (id is not null && target is not null)
                result[id] = (type ?? "", target);
        }

        return result;
    }

    /// <summary>
    /// 流式枚举 workbook.xml 中的 &lt;sheet&gt; 元素（保持原始顺序），
    /// 每项为 (name, rId)；调用方自行查 rels 做 rId → 部件解析。
    /// </summary>
    public static List<(string Name, string RId)> EnumerateSheets(ZipArchive zip)
    {
        var result = new List<(string, string)>();
        var wbEntry = zip.GetEntry("xl/workbook.xml");
        if (wbEntry == null) return result;

        using var stream = wbEntry.Open();
        using var reader = XlsxXml.CreatePartReader(stream, XlsxXml.BufferSmall);

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "sheet") continue;

            string? name = reader.GetAttribute("name");
            string? rid = reader.GetAttribute("id", NsRelationships);
            if (name is { Length: > 0 } && rid is not null)
                result.Add((name, rid));
        }

        return result;
    }
}
