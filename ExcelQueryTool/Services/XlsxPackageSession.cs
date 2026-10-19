using System.IO;
using System.IO.Compression;

namespace ExcelQueryTool.Services;

/// <summary>
/// 一个 xlsx 文件的「打开会话」：包只开一次，共享部件只解析一次。
///
/// 存在的理由（2026-09-23 实测）：
/// 此前每次 <c>ReadLayoutWithOccupancy</c> 都自己 <c>new FileStream</c> + <c>new ZipArchive</c>，
/// 而 ZipArchive 构造必须读完整个中央目录。14.8 MB / 268 条目的包实测光这一步
/// <b>4.41 ms/次</b>，占单表总成本（7.21 ms）的 <b>61%</b>；sharedStrings / styles 也每次重解析。
/// 于是「切表 / 搜索 / 勾选显示图片重载」每次都白交这笔钱，且随表数放大
/// （268 表累计 1.9 s，其中 1.2 s 纯开包）。
///
/// 改用会话后：同一个文件的开包 + 部件解析在整个生命周期内只做一次，
/// 单表成本降到「净扫描」的 2.8 ms，且不再随表数增长。
///
/// 线程安全：ZipArchive 本身不支持并发读，因此本类用锁把「打开部件流」串行化。
/// 调用方须保证同一会话不被两个线程同时用于不同工作表的长时扫描
/// （当前流程天然满足：扫描串行发生在后台线程，图片索引在另一时刻）。
/// </summary>
internal sealed class XlsxPackageSession : IDisposable
{
    private readonly string _filePath;
    private readonly object _gate = new();

    private FileStream? _stream;
    private ZipArchive? _zip;
    private bool _opened;
    private bool _disposed;

    private List<string>? _sharedStrings;
    private HashSet<int>? _dateStyles;
    private Dictionary<string, string?> _sheetPartCache = new(StringComparer.Ordinal);
    private HashSet<string> _missingSheetParts = new(StringComparer.Ordinal);

    public XlsxPackageSession(string filePath) => _filePath = filePath;

    /// <summary>工作表名 → 包内部件路径（结果按表名缓存，重复查表不再走 workbook/rels 解析）</summary>
    public string? LocateSheet(string worksheetName)
    {
        lock (_gate)
        {
            if (_sheetPartCache.TryGetValue(worksheetName, out var cached)) return cached;
            if (_missingSheetParts.Contains(worksheetName)) return null;
        }

        EnsureOpened();

        string? part;
        lock (_gate)
        {
            part = _zip == null ? null : XlsxPackageLocator.Locate(_zip, worksheetName);
            if (part == null) _missingSheetParts.Add(worksheetName);
            else _sheetPartCache[worksheetName] = part;
        }
        return part;
    }

    /// <summary>共享部件（sharedStrings / 日期样式）：打开时解析一次，失败退化为空</summary>
    public (List<string> SharedStrings, HashSet<int> DateStyles) SharedParts
    {
        get
        {
            EnsureOpened();
            lock (_gate)
            {
                return (_sharedStrings ?? [], _dateStyles ?? []);
            }
        }
    }

    /// <summary>
    /// 在会话锁内执行一段「独占使用 zip」的代码。
    ///
    /// 为什么需要：扫描（ScanRows / ProbeMaxColumn / LoadColumnWidths）会持续从
    /// <see cref="ZipArchive"/> 打开部件流，而图片索引构建线程也会读同一个包。
    /// ZipArchive 不支持并发读，因此整段扫描必须串行化。
    /// 锁粒度取「一次完整扫描」而非「单次部件读取」：扫描内部会多次开流，
    /// 逐次加锁既无收益（扫描本身不可打断）又易漏。
    /// </summary>
    public T RunLocked<T>(Func<ZipArchive, T> body, T fallback)
    {
        EnsureOpened();

        lock (_gate)
        {
            if (_zip == null) return fallback;
            return body(_zip);
        }
    }

    private void EnsureOpened()
    {
        lock (_gate)
        {
            if (_opened || _disposed) return;
            _opened = true;
            try
            {
                _stream = new FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                _zip = new ZipArchive(_stream, ZipArchiveMode.Read);
                // 复用 XlsxSheetReader 的唯一一份实现（避免副本漂移：
                // 曾经照抄的副本把「cellXfs 下标」错写成「numFmtId」，导致日期全部显示成序列号）
                _sharedStrings = XlsxSheetReader.LoadSharedStrings(_zip);
                _dateStyles = XlsxSheetReader.LoadDateStyles(_zip);
            }
            catch
            {
                _zip = null;   // 打开失败：后续一律退化为「空表」，不抛
            }
        }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed) return;
            _disposed = true;
            _zip?.Dispose();
            _stream?.Dispose();
            _zip = null;
            _stream = null;
            _sheetPartCache.Clear();
            _missingSheetParts.Clear();
            _sharedStrings = null;
            _dateStyles = null;
        }
    }
}

/// <summary>把工作表名解析为包内部件路径。</summary>
internal static class XlsxPackageLocator
{
    public static string? Locate(ZipArchive zip, string worksheetName)
    {
        var rels = XlsxRelationships.ReadRels(zip, "xl/workbook.xml");

        foreach (var (name, rid) in XlsxRelationships.EnumerateSheets(zip))
        {
            if (!string.Equals(name, worksheetName, StringComparison.Ordinal)) continue;

            if (rels.TryGetValue(rid, out var rel) && rel.Type == XlsxRelationships.RelTypeWorksheet)
                return XlsxPartPath.ResolvePartPath("xl/workbook.xml", rel.Target);
        }

        return null;
    }
}

/// <summary>
/// OOXML 包内部件路径的解析工具。
/// 
/// 原先 <see cref="XlsxPackageLocator"/> 与 <see cref="XlsxPictureIndex"/>
/// 各持一套逐字符等价的实现（BuildRelsPath / ResolvePartPath / CombineRelative / NormalizePath）。
/// 包内路径规则一旦需要调整（例如 TargetMode="External" 的处理），双份实现必然漏改 ——
/// 现统一收敛到此处。
/// </summary>
internal static class XlsxPartPath
{
    /// <summary>xl/worksheets/sheet1.xml → xl/worksheets/_rels/sheet1.xml.rels</summary>
    public static string BuildRelsPath(string partPath)
    {
        int slash = partPath.LastIndexOf('/');
        string dir = slash >= 0 ? partPath[..slash] : "";
        string file = slash >= 0 ? partPath[(slash + 1)..] : partPath;
        return string.IsNullOrEmpty(dir) ? $"_rels/{file}.rels" : $"{dir}/_rels/{file}.rels";
    }

    /// <summary>把 rels 里的相对 Target 解析为包内绝对路径，并归一化 . / ..</summary>
    public static string? ResolvePartPath(string basePartPath, string target)
    {
        if (string.IsNullOrEmpty(target)) return null;

        // 以 / 开头表示包根绝对路径
        string combined = target.StartsWith('/')
            ? target.TrimStart('/')
            : CombineRelative(basePartPath, target);

        return NormalizePath(combined);
    }

    private static string CombineRelative(string basePartPath, string target)
    {
        int slash = basePartPath.LastIndexOf('/');
        string dir = slash >= 0 ? basePartPath[..slash] : "";
        return string.IsNullOrEmpty(dir) ? target : $"{dir}/{target}";
    }

    private static string? NormalizePath(string path)
    {
        var segments = path.Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries);
        var stack = new List<string>(segments.Length);

        foreach (var seg in segments)
        {
            if (seg == ".") continue;
            if (seg == "..")
            {
                if (stack.Count == 0) return null;
                stack.RemoveAt(stack.Count - 1);
            }
            else stack.Add(seg);
        }

        return stack.Count == 0 ? null : string.Join('/', stack);
    }
}
