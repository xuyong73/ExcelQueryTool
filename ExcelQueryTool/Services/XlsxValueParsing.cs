using System.Globalization;

namespace ExcelQueryTool.Services;

/// <summary>
/// 单元格值解析的共享实现（列号字母、手写非负整数、&lt;v&gt; 原文 → 值、日期序列号）。
///
/// 两条扫描路径（<see cref="XlsxSheetReader"/> 的 XmlReader 路径与 <see cref="XlsxByteScanner"/>
/// 的字节路径）必须产出逐值一致的结果。历史上这两份规则靠「镜像拷贝 + 注释」同步，
/// 曾因副本把 cellXfs 下标错写成 numFmtId 导致日期全部显示成序列号；
/// 本类是两条路径共用的唯一一份实现，字节侧保留 <c>TryFastNumericPassthrough</c> 作可选快路径前端。
/// </summary>
internal static class XlsxValueParsing
{
    /// <summary>
    /// <c>r="C12"</c> 的前导列字母 → 列号（"C"→3，"AA"→27）。
    /// 直接按字符累加，不做 <c>[..ci].ToUpperInvariant()</c> 切片（避免每个单元格的临时字符串分配）。
    /// 大小写同时接受（OOXML 规范要求大写，但小写也不应解析失败）。
    /// 返回 0 表示无前导字母（调用方回退到顺序列号）。
    /// </summary>
    public static int ParseColumnLetters(string s)
    {
        int result = 0;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c >= 'A' && c <= 'Z') result = result * 26 + (c - 'A' + 1);
            else if (c >= 'a' && c <= 'z') result = result * 26 + (c - 'a' + 1);
            else break;     // 到达数字部分（行号），结束
        }
        return result;
    }

    /// <summary><paramref name="s"/> 的前导列字母 → 列号（OOXML 引用为纯 ASCII，逐字节处理）。</summary>
    public static int ParseColumnLettersAscii(ReadOnlySpan<byte> s)
    {
        int result = 0;
        for (int k = 0; k < s.Length; k++)
        {
            byte c = s[k];
            if (c >= (byte)'A' && c <= (byte)'Z') result = result * 26 + (c - (byte)'A' + 1);
            else if (c >= (byte)'a' && c <= (byte)'z') result = result * 26 + (c - (byte)'a' + 1);
            else break;
        }
        return result;
    }

    /// <summary>
    /// 手写非负整数解析（样式索引、共享字符串索引、行号等）。
    /// 相比 <c>int.TryParse(s, NumberStyles.Integer, InvariantCulture, out _)</c>：
    /// 无需检查符号/空白/千分位、不构造 NumberFormatInfo 状态机。热路径上每个单元格
    /// 至少调用一次，收益可观。溢出时返回 false（与原 TryParse 行为一致）。
    /// </summary>
    public static bool TryParseNonNegativeInt(string s, out int value)
    {
        value = 0;
        if (s.Length == 0) return false;

        int acc = 0;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c < '0' || c > '9') return false;
            int digit = c - '0';
            // 与 int.TryParse 一致：溢出即失败，不返回截断值
            if (acc > (int.MaxValue - digit) / 10) { value = 0; return false; }
            acc = acc * 10 + digit;
        }

        value = acc;
        return true;
    }

    /// <summary>字节版非负整数解析：OOXML 属性值恒为 ASCII 数字串，逐字节处理零分配。</summary>
    public static bool TryParseNonNegativeInt(ReadOnlySpan<byte> s, out int value)
    {
        value = 0;
        if (s.Length == 0) return false;

        int acc = 0;
        for (int k = 0; k < s.Length; k++)
        {
            byte c = s[k];
            if (c < (byte)'0' || c > (byte)'9') return false;
            int digit = c - (byte)'0';
            if (acc > (int.MaxValue - digit) / 10) { value = 0; return false; }
            acc = acc * 10 + digit;
        }
        value = acc;
        return true;
    }

    /// <summary>
    /// &lt;v&gt; 原文 → 单元格值（XmlReader 路径入口，<paramref name="cellType"/> 为 <c>t</c> 属性原文）。
    /// 共享字符串（t="s"）按索引查表，日期样式（s 命中日期 numFmt）序列号转 DateTime。
    /// </summary>
    public static object? ResolveCellType(string raw, string? cellType, bool isDate, List<string> sharedStrings)
    {
        if (cellType == "s")
        {
            // 共享字符串索引用同一个手写解析器（原实现走通用 int.TryParse）
            return TryParseNonNegativeInt(raw, out int si)
                   && si < sharedStrings.Count ? sharedStrings[si] : raw;
        }

        // 注：t="s" 已在上面提前返回；这里 str/e/日期/数值 与字节侧 kind 2/3/0 的分支一一对应
        switch (cellType)
        {
            case "str":
            case "e":
                return raw;
            case "b":
                return raw == "1";
            default:
                return ConvertNumericOrDate(raw, isDate, passthrough: null);
        }
    }

    /// <summary>
    /// 字节路径入口：<paramref name="cellTypeKind"/> 0=无t/数字 1=s 2=str 3=e 4=b 5=inlineStr 6=其他。
    /// <paramref name="fastNumericPassthrough"/> 为字节侧的 G15 原样快路径（返回 null = 不适用，仅作用于数字分支）。
    /// </summary>
    public static object? ResolveCellTypeKind(string raw, int cellTypeKind, bool isDate, List<string> sharedStrings,
        Func<string, string?>? fastNumericPassthrough = null)
    {
        switch (cellTypeKind)
        {
            case 1:     // t="s"：共享字符串索引
                return TryParseNonNegativeInt(raw, out int si) && si < sharedStrings.Count
                    ? sharedStrings[si]
                    : raw;
            case 2:     // t="str"
            case 3:     // t="e"
                return raw;
            case 4:     // t="b"
                return raw == "1";
            default:    // 无 t / 其他：日期或数字
                return ConvertNumericOrDate(raw, isDate, fastNumericPassthrough);
        }
    }

    /// <summary>日期序列号 / 数值分支（两条路径共用）。</summary>
    private static object? ConvertNumericOrDate(string raw, bool isDate, Func<string, string?>? passthrough)
    {
        // 日期样式：序列号 → DateTime（Excel 1900 日期系统，含闰年 bug 偏移）
        if (isDate && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double serial))
            return FromOleAutomationDate(serial);

        // 数字规整快路径（仅字节侧提供）
        if (passthrough is not null && passthrough(raw) is { } direct)
            return direct;

        // 普通数值：xml 里的 raw 可能带二进制浮点噪声（如 0.1 存成
        // 0.10000000000000001，13 位小数）。Excel 只保留 15 位有效数字显示，
        // 这里按同口径规整，与 Excel 界面显示一致。
        if (!isDate && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double num))
            return num.ToString("G15", CultureInfo.InvariantCulture);
        return raw;
    }

    /// <summary>
    /// Excel 日期序列号 → DateTime。
    ///
    /// .NET 的 DateTime.FromOADate 以 1899-12-30 为基准，正好吸收 Excel
    /// 「1900 年当闰年」的历史错误，序列号 ≥ 61 时结果与 Excel 显示一致。
    /// 序列号 &lt; 61 落在该错误区间之前，需回退一天。
    /// </summary>
    public static DateTime FromOleAutomationDate(double serial)
    {
        var value = DateTime.FromOADate(serial);
        return serial < 61 ? value.AddDays(-1) : value;
    }
}

/// <summary>单元格值 → 显示字符串。全项目唯一的一份实现（ExcelDataManager 亦委托于此）。</summary>
internal static class XlsxCellValue
{
    public static string ToDisplayString(object? value) => value switch
    {
        null or DBNull => string.Empty,
        string s => s,
        DateTime dt => dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        bool b => b ? "TRUE" : "FALSE",
        double d => d.ToString(CultureInfo.InvariantCulture),
        _ => value.ToString() is { Length: > 0 } str ? str : string.Empty,
    };
}

/// <summary>
/// 列命名规则的共享实现：空表头 → &lt;列N&gt;（N = 源列号），重名追加 _2/_3（首次出现不加后缀）。
///
/// xlsx 路径（<see cref="ExcelDataManager.GetColumnMetadataWithLayout"/>）与文本路径
/// （<see cref="TextTableReader.BuildColumns"/>）必须产出完全一致的列名，否则表头与行数据
/// 脱节——本类是唯一的一份实现。
/// </summary>
internal sealed class HeaderNormalizer
{
    private readonly Dictionary<string, int> _usedNames = new(StringComparer.Ordinal);

    /// <summary>
    /// 归一化一列的表头，返回最终列名。
    /// <paramref name="sourceCol"/> 为 1 基源列号（用于 &lt;列N&gt; 的 N）；
    /// <paramref name="rawHeader"/> 为原始表头（调用方决定是否 Trim）。
    /// </summary>
    public string Normalize(int sourceCol, string? rawHeader)
    {
        string header = rawHeader ?? "";

        // 保留的空表头列统一命名为 <列N>（N 为源列号）
        if (string.IsNullOrWhiteSpace(header)) header = $"<列{sourceCol}>";

        // 重名追加 _2 / _3 ...（同名首次不追加后缀）
        if (_usedNames.TryGetValue(header, out int seen))
        {
            seen++;
            _usedNames[header] = seen;
            return $"{header}_{seen}";
        }

        _usedNames[header] = 1;
        return header;
    }
}

/// <summary>
/// 列级「数字列」判定的共享口径：抽样 <see cref="SampleRows"/> 行数据，
/// 非空值中数字占比 ≥ <see cref="RatioPercent"/>/100 且样本 ≥ <see cref="MinSample"/> 判为数字列。
///
/// xlsx 路径（<see cref="XlsxSheetReader"/>）与文本路径（<see cref="TextTableReader"/>）
/// 必须同口径，否则两种格式的列对齐观感不一致；本类是唯一一份实现
/// （历史上两边各持一份 0.8 / 3 / 200 常量，靠注释「与 xlsx 侧同口径」人工同步）。
/// </summary>
internal static class NumericColumnClassifier
{
    /// <summary>抽样行数上限（取数据区前 N 行，避免为对齐多扫一遍全表）</summary>
    public const int SampleRows = 200;

    /// <summary>样本下限：非空值少于此数不下结论（避免前几行恰好是数字就定成数字列）</summary>
    public const int MinSample = 3;

    /// <summary>数字占比阈值（百分数）：非空值中数字占比 ≥ 0.8</summary>
    public const int RatioPercent = 80;

    /// <summary>单列判定：<paramref name="hits"/> = 数字值数，<paramref name="total"/> = 非空值数。</summary>
    public static bool Classify(int hits, int total)
        => total >= MinSample && hits * 100 >= total * RatioPercent;
}
