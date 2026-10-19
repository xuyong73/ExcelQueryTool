using System.Globalization;
using System.Text;

namespace ExcelQueryTool.Services;

/// <summary>
/// 手写字节级 sheetData 扫描器。
///
/// 为什么不用 <see cref="System.Xml.XmlReader"/>：
/// 百万行 × 6 列实测，XmlReader 解析段 3436 ms，其中约 52% 花在
/// <c>MoveToAttribute</c> / <c>Value</c> 的字符串物化上（3775 万个节点，
/// 每次属性访问都要切出一个 string）。本扫描器直接在 UTF-8 字节缓冲上定位
/// 标签与属性名，只在「确实要落格」时才解码成 string，同文件解析段降到 625 ms。
///
/// 设计约束：
/// - <b>只负责 sheetData 内容</b>。<c>&lt;cols&gt;</c> / <c>&lt;mergeCells&gt;</c> /
///   <c>&lt;dimension&gt;</c> 等结构仍由 <c>XmlReader</c> 读取（调用方保留原逻辑），
///   本类从 <c>&lt;sheetData&gt;</c> 起扫到其结束标签为止。
/// - <b>语义与 <see cref="XlsxSheetReader"/> 的 XmlReader 扫码路径等价</b>：
///   行号（含缺 <c>r</c> 的顺序号）、列号（含回退顺序列）、单元格类型、
///   日期样式、共享字符串、内联字符串、<c>vm</c>（richValue）抑制、
///   空单元格 <c>NoteColumnSeen</c>，逐项镜像。
/// - 复用调用方传入的 <see cref="XlsxRowBuffer"/>，跨行零分配。
/// </summary>
internal static class XlsxByteScanner
{
    private const byte Lt = (byte)'<';
    private const byte Gt = (byte)'>';
    private const byte Slash = (byte)'/';
    private const byte Eq = (byte)'=';
    private const byte DQuote = (byte)'"';
    private const byte SQuote = (byte)'\'';

    private static readonly UTF8Encoding Utf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    private const int StackDecodeLimit = 1024;

    /// <summary>
    /// 扫 <paramref name="utf8"/> 里的 sheetData 行，逐行回调 <paramref name="visitor"/>。
    /// 返回 false 表示不支持该输入（调用方回退到 XmlReader 路径）。
    /// <paramref name="endOffset"/>：扫描停止处的绝对偏移（正常扫完 = 长度；
    /// visitor 提前停止 = 最后一个已投出行结束标签之后）。
    /// </summary>
    public static bool TryScan(
        ReadOnlySpan<byte> utf8,
        List<string> sharedStrings,
        HashSet<int> dateStyles,
        CancellationToken token,
        RowVisitorLike visitor,
        out int endOffset)
    {
        endOffset = utf8.Length;
        int n = utf8.Length;
        int i = 0;

        // BOM：UTF-8 跳过；UTF-16/32 不支持 → 交回 XmlReader
        if (n >= 3 && utf8[0] == 0xEF && utf8[1] == 0xBB && utf8[2] == 0xBF) i = 3;
        else if (n >= 2 && utf8[0] == 0xFF && utf8[1] == 0xFE) return false;
        else if (n >= 2 && utf8[0] == 0xFE && utf8[1] == 0xFF) return false;
        else if (n >= 4 && utf8[0] == 0x00 && utf8[1] == 0x00) return false;

        int sd = IndexOfOpenTag(utf8, i, "sheetData");
        if (sd < 0) return false;

        int p = IndexOfByte(utf8, sd, Gt);
        if (p < 0) return false;
        p++;

        bool ok = ScanRowsCore(utf8.Slice(p), 0, sharedStrings, dateStyles, token, visitor, out int pos);
        endOffset = p + pos;
        return ok;
    }

    /// <summary>
    /// 扫一个「行对齐」的连续区间（并行尾巴的 worker 入口）。
    /// <paramref name="slice"/> 必须从一个 &lt;row 开标签开始、到下一个 &lt;row 之前结束；
    /// <paramref name="seedSequentialRow"/> = 该块之前已出现过的 row 元素总数 ——
    /// 缺 <c>r</c> 属性的行按顺序号兜底时才能得到与全量扫描一致的行号
    /// （顺序号语义 = 「第几个 row 元素」，不是物理行号差）。
    /// 与 <see cref="TryScan"/> 共用同一份核心循环，杜绝两份解析逻辑漂移。
    /// </summary>
    public static void ScanChunk(
        ReadOnlySpan<byte> slice,
        int seedSequentialRow,
        List<string> sharedStrings,
        HashSet<int> dateStyles,
        CancellationToken token,
        RowVisitorLike visitor)
    {
        ScanRowsCore(slice, seedSequentialRow, sharedStrings, dateStyles, token, visitor, out _);
    }

    /// <summary>
    /// 行扫描核心循环（单份实现）。输入切片从首个待处理标签开始；
    /// 全量模式传入 sheetData 起始，分块模式传入行边界。
    /// </summary>
    private static bool ScanRowsCore(
        ReadOnlySpan<byte> utf8,
        int seedSequentialRow,
        List<string> sharedStrings,
        HashSet<int> dateStyles,
        CancellationToken token,
        RowVisitorLike visitor,
        out int position)
    {
        int n = utf8.Length;
        int p = 0;

        var row = new XlsxRowBuffer();

        // 行/列游标
        int currentColumn = 1;
        int sequentialColumn = 1;
        int sequentialRow = seedSequentialRow;
        int currentRowNumber = 0;
        bool rowOpen = false;

        // 单元格状态：唯一收敛点是 CommitCell（镜像 XmlReader 版的 EndElement: c）
        int cellTypeKind = 0;      // 0=无t/数字 1=s 2=str 3=e 4=b 5=inlineStr 6=其他
        bool cellIsDate = false;
        int cellVm = 0;
        bool inValue = false;      // 在 <v> 内
        bool inInlineStr = false;  // <c t="inlineStr">，等待 <is> 内 <t>
        int vStart = -1;
        int tStart = -1;
        int tEnd = -1;

        long counter = 0;

        while (p < n)
        {
            if ((++counter & 0x3FF) == 0) token.ThrowIfCancellationRequested();

            int lt = IndexOfByte(utf8, p, Lt);
            if (lt < 0) break;

            bool closing = lt + 1 < n && utf8[lt + 1] == Slash;
            int nameStart = closing ? lt + 2 : lt + 1;
            int nameEnd = ScanName(utf8, nameStart);
            int g = IndexOfByte(utf8, nameEnd, Gt);
            if (g < 0) break;

            bool selfClosing = !closing && g > nameStart && utf8[g - 1] == Slash;
            int attrEnd = selfClosing ? g - 1 : g;
            var local = LocalName(utf8, nameStart, nameEnd);

            if (closing)
            {
                if (SeqEqual(local, "v"))
                {
                    // <v> 文本结束 → 落格。镜像生产语义：XmlReader 对 <v></v>（空）
                    // 与 <v>纯空白</v> 都不产生 Text 节点（IgnoreWhitespace）→ 不落格；
                    // vm（richValue）单元格的值恒为 #VALUE!，也不落格。
                    if (inValue && vStart >= 0 && cellVm == 0 && lt > vStart && !IsAllXmlWhitespace(utf8, vStart, lt))
                    {
                        // 与 XmlReader 路径共用同一份值解析；快路径前端（FastNumericPassthrough）仅作用于数字分支
                        row.Add(currentColumn,
                                XlsxValueParsing.ResolveCellTypeKind(DecodeEntities(DecodeUtf8(utf8, vStart, lt)), cellTypeKind, cellIsDate, sharedStrings, FastNumericPassthrough));
                        sequentialColumn = currentColumn + 1;
                    }
                    inValue = false;
                    vStart = -1;
                }
                else if (SeqEqual(local, "t"))
                {
                    if (inInlineStr && tStart >= 0)
                    {
                        tEnd = lt;              // 文本到 </t> 为止
                        inInlineStr = false;    // 镜像生产：ReadElementContentAsString 消费第一个 <t> 后即关闭，后续 <t> 忽略
                    }
                }
                else if (SeqEqual(local, "c"))
                {
                    // 内联文本落格（生产在 ReadElementContentAsString 时消费、</c> 时落格）：
                    // 区间 [tStart, tEnd)，空串不落格（inlineText.Length > 0 才落），空白保留。
                    if (tEnd > tStart)
                    {
                        row.Add(currentColumn, DecodeEntities(DecodeUtf8(utf8, tStart, tEnd)));
                        sequentialColumn = currentColumn + 1;
                    }
                    CommitCell();
                }
                else if (SeqEqual(local, "row"))
                {
                    if (rowOpen)
                    {
                        rowOpen = false;
                        if (!visitor(currentRowNumber, row))
                        {
                            // 提前停止：position 落在本行结束标签之后（下一行起点），
                            // 保证头部阶段截停时尾部从「未投出的第一行」开始，绝不丢行
                            position = g + 1;
                            return true;
                        }
                    }
                }
                else if (SeqEqual(local, "sheetData"))
                {
                    position = g + 1;
                    return true;
                }

                p = g + 1;
                continue;
            }

            // ── 开标签 ──
            if (SeqEqual(local, "row"))
            {
                // 镜像生产：XmlReader 对 <row r="1"/>（自闭合）只产生 Element 事件、
                // 不产生 EndElement → 该行从不投出，缓冲被下一行 Reset 覆盖。
                // 因此这里不做「上一行未闭合就补投」——丢失行为与生产一致。
                row.Reset();
                sequentialColumn = 1;
                sequentialRow++;
                currentRowNumber = ReadAttrInt(utf8, nameEnd, attrEnd, "r", sequentialRow);
                rowOpen = !selfClosing;
            }
            else if (SeqEqual(local, "c"))
            {
                CommitCell();       // 清上一格残留（正常路径已在 </c> 清过）
                cellTypeKind = 0;
                cellIsDate = false;
                cellVm = 0;

                // 镜像生产的列号语义（<c> 分支）：
                // - 完全无属性 → currentColumn = sequentialColumn
                // - 有属性但无 r → currentColumn 保持上一格的值（不是顺序列！）
                // - 有 r → col > 0 ? col : sequentialColumn
                if (SkipSpace(utf8, nameEnd, attrEnd) >= attrEnd)
                    currentColumn = sequentialColumn;

                if (selfClosing)
                {
                    // <c r="I1"/>：无值，但必须记「本行铺到第 I 列」
                    ApplyCellAttributes(utf8, nameEnd, attrEnd, dateStyles, ref cellTypeKind, ref cellIsDate, ref cellVm, ref currentColumn, ref sequentialColumn);
                    row.NoteColumnSeen(currentColumn);
                    CommitCell();
                }
                else
                {
                    ApplyCellAttributes(utf8, nameEnd, attrEnd, dateStyles, ref cellTypeKind, ref cellIsDate, ref cellVm, ref currentColumn, ref sequentialColumn);
                    row.NoteColumnSeen(currentColumn);
                    if (cellTypeKind == 5) inInlineStr = true;
                }
            }
            else if (SeqEqual(local, "v"))
            {
                if (selfClosing)
                {
                    inValue = false;
                    vStart = -1;
                }
                else
                {
                    inValue = true;
                    vStart = g + 1;
                }
            }
            else if (SeqEqual(local, "is"))
            {
                if (selfClosing) inInlineStr = false;
            }
            else if (SeqEqual(local, "t"))
            {
                if (inInlineStr && !selfClosing)
                {
                    tStart = g + 1;
                    tEnd = -1;
                }
            }

            p = g + 1;
        }

        // 镜像生产：XML 截断时 XmlReader 的 Read() 返回 false，未闭合的行同样丢失。
        position = n;
        return true;

        void CommitCell()
        {
            inValue = false;
            inInlineStr = false;
            vStart = -1;
            tStart = -1;
            tEnd = -1;
        }
    }

    // ───────────────────────── 并行切块 ─────────────────────────

    /// <summary>
    /// 在 [from, to) 内找所有 &lt;row 开标签的起始偏移。
    /// 用于把尾部字节区间按行边界切成并行块。
    ///
    /// 为什么安全：裸 <c>&lt;</c> 在 XML 属性值与文本内容里都是非法的（必须写成 &amp;lt;），
    /// 因此按字节序列找 <c>&lt;row</c> 再核对名字终止符不会误匹配数据内容；
    /// <c>&lt;rowBreaks</c> 一类前缀重名靠「终止符核对」排除。
    ///
    /// 已知边界：带命名空间前缀的行标签（<c>&lt;x:row&gt;</c>）不被本方法识别 ——
    /// 影响仅限「该类文件 + 缺 r 属性行」同时出现时顺序号种子偏差（Excel 本身从不写前缀行标签，
    /// 全前缀文件会退化为单块串行扫描，行号仍正确）。
    /// </summary>
    public static List<int> FindRowStarts(ReadOnlySpan<byte> b, int from, int to)
    {
        var result = new List<int>();
        ReadOnlySpan<byte> pattern = "<row"u8;
        int k = from;
        while (k < to)
        {
            int hit = b.Slice(k, to - k).IndexOf(pattern);   // 向量化序列搜索：整段 SIMD 扫 + 仅百万级命中
            if (hit < 0) break;
            k += hit;

            // 命中后核对名字终止符：排除 <rowBreaks 一类前缀重名（其 'B' 不是名字终止符）
            int nameEnd = k + pattern.Length;
            if (nameEnd < to && IsNameEnd(b[nameEnd]))
                result.Add(k);

            k += pattern.Length;
        }
        return result;
    }

    // ───────────────────────── 属性解析 ─────────────────────────

    /// <summary>
    /// 一趟遍历 <c>&lt;c&gt;</c> 的属性（顺序无保证 → 按名字分派）。
    /// 列号语义与生产一致：r 存在时 <c>col &gt; 0 ? col : sequentialColumn</c>；
    /// r 不存在时不改动（由调用方决定无属性时的回退）。
    /// </summary>
    private static void ApplyCellAttributes(
        ReadOnlySpan<byte> b, int from, int attrEnd,
        HashSet<int> dateStyles,
        ref int cellTypeKind, ref bool cellIsDate, ref int cellVm,
        ref int currentColumn, ref int sequentialColumn)
    {
        int a = from;

        while (true)
        {
            a = SkipSpace(b, a, attrEnd);
            if (a >= attrEnd) break;

            int anStart = a;
            int anEnd = ScanName(b, anStart);
            int eq = SkipSpace(b, anEnd, attrEnd);
            if (eq >= attrEnd || b[eq] != Eq) { a = anEnd + 1; continue; }

            int q = SkipSpace(b, eq + 1, attrEnd);
            if (q >= attrEnd) break;

            byte quote = b[q];
            if (quote != DQuote && quote != SQuote) { a = q + 1; continue; }

            int vs = q + 1;
            int ve = IndexOfByte(b, vs, quote);
            if (ve < 0 || ve > attrEnd) break;

            var aname = LocalName(b, anStart, anEnd);
            var aval = b.Slice(vs, ve - vs);

            if (SeqEqual(aname, "t"))
            {
                if (SeqEqual(aval, "s")) cellTypeKind = 1;
                else if (SeqEqual(aval, "str")) cellTypeKind = 2;
                else if (SeqEqual(aval, "e")) cellTypeKind = 3;
                else if (SeqEqual(aval, "b")) cellTypeKind = 4;
                else if (SeqEqual(aval, "inlineStr")) cellTypeKind = 5;
                else cellTypeKind = 6;
            }
            else if (SeqEqual(aname, "vm"))
            {
                // vm="N"：新版 Excel 内嵌图片（richValue）元数据索引，真实内容是图片
                if (XlsxValueParsing.TryParseNonNegativeInt(aval, out int vmIdx)) cellVm = vmIdx + 1;   // 1 基；0 = 无
            }
            else if (SeqEqual(aname, "s"))
            {
                if (XlsxValueParsing.TryParseNonNegativeInt(aval, out int styleIdx)) cellIsDate = dateStyles.Contains(styleIdx);
            }
            else if (SeqEqual(aname, "r"))
            {
                int col = XlsxValueParsing.ParseColumnLettersAscii(aval);
                currentColumn = col > 0 ? col : sequentialColumn;
            }

            a = ve + 1;
        }
    }

    // ───────────────────────── 基础工具 ─────────────────────────

    private static int IndexOfByte(ReadOnlySpan<byte> b, int from, byte target)
    {
        for (int k = from; k < b.Length; k++)
            if (b[k] == target) return k;
        return -1;
    }

    private static int SkipSpace(ReadOnlySpan<byte> b, int from, int limit)
    {
        while (from < limit)
        {
            byte c = b[from];
            if (c != (byte)' ' && c != (byte)'\t' && c != (byte)'\r' && c != (byte)'\n') break;
            from++;
        }
        return from;
    }

    /// <summary>区间是否全为 XML 空白（空格/制表/回车/换行）。用于镜像 XmlReader 的 IgnoreWhitespace：&lt;v&gt;&lt;/v&gt; 与 &lt;v&gt;纯空白&lt;/v&gt; 都不产生 Text 节点。</summary>
    private static bool IsAllXmlWhitespace(ReadOnlySpan<byte> b, int start, int end)
    {
        for (int k = start; k < end; k++)
        {
            byte c = b[k];
            if (c != (byte)' ' && c != (byte)'\t' && c != (byte)'\r' && c != (byte)'\n') return false;
        }
        return true;
    }

    private static bool IsNameEnd(byte c) =>
        c == (byte)' ' || c == (byte)'\t' || c == (byte)'\r' || c == (byte)'\n' ||
        c == Gt || c == Slash || c == Eq;

    private static int ScanName(ReadOnlySpan<byte> b, int from)
    {
        int k = from;
        while (k < b.Length && !IsNameEnd(b[k])) k++;
        return k;
    }

    /// <summary>取 local name（去掉 <c>ns:</c> 前缀）</summary>
    private static ReadOnlySpan<byte> LocalName(ReadOnlySpan<byte> b, int ns, int ne)
    {
        int local = ns;
        for (int k = ns; k < ne; k++)
            if (b[k] == (byte)':') local = k + 1;
        return b.Slice(local, ne - local);
    }

    private static bool SeqEqual(ReadOnlySpan<byte> b, string s)
    {
        if (b.Length != s.Length) return false;
        for (int k = 0; k < b.Length; k++)
            if (b[k] != (byte)s[k]) return false;
        return true;
    }

    private static int IndexOfOpenTag(ReadOnlySpan<byte> b, int from, string name)
    {
        int k = from;
        while (true)
        {
            k = IndexOfByte(b, k, Lt);
            if (k < 0) return -1;

            int ns = k + 1;
            if (ns < b.Length && b[ns] == Slash) { k++; continue; }     // 结束标签

            int ne = ScanName(b, ns);
            if (ne <= ns) { k++; continue; }

            if (SeqEqual(LocalName(b, ns, ne), name)) return k;
            k++;
        }
    }

    private static int ReadAttrInt(ReadOnlySpan<byte> b, int from, int attrEnd, string name, int fallback)
    {
        int a = from;
        while (true)
        {
            a = SkipSpace(b, a, attrEnd);
            if (a >= attrEnd) return fallback;

            int anStart = a;
            int anEnd = ScanName(b, anStart);
            int eq = SkipSpace(b, anEnd, attrEnd);
            if (eq >= attrEnd || b[eq] != Eq) { a = anEnd + 1; continue; }

            int q = SkipSpace(b, eq + 1, attrEnd);
            if (q >= attrEnd) return fallback;

            byte quote = b[q];
            if (quote != DQuote && quote != SQuote) { a = q + 1; continue; }

            int vs = q + 1;
            int ve = IndexOfByte(b, vs, quote);
            if (ve < 0 || ve > attrEnd) return fallback;

            if (SeqEqual(LocalName(b, anStart, anEnd), name))
                return XlsxValueParsing.TryParseNonNegativeInt(b.Slice(vs, ve - vs), out int val) ? val : fallback;

            a = ve + 1;
        }
    }

    private static string DecodeUtf8(ReadOnlySpan<byte> b, int start, int end)
    {
        if (end <= start) return string.Empty;

        int len = end - start;
        if (len <= StackDecodeLimit)
        {
            Span<char> buf = stackalloc char[StackDecodeLimit];
            int written = Utf8.GetChars(b.Slice(start, len), buf);
            return new string(buf[..written]);
        }
        return Utf8.GetString(b.Slice(start, len));
    }

    /// <summary>
    /// XML 实体解码，镜像 XmlReader 对字符数据的解码行为（预定义实体 + 字符引用）。
    /// Excel 会把公式结果（t="str"）与内联字符串里的非 ASCII 字符写成 <c>&amp;#28450;</c>
    /// 形态的字符引用，不解码则界面显示的就是原样引用文本。
    /// 绝大多数值不含 '&'，一次快速检查即零成本返回。
    /// 不支持的实体形态（未定义实体、非法/越界引用）抛 <see cref="System.Xml.XmlException"/>——
    /// 与 XmlReader 在同样输入下的行为一致；绝不能返回失败标记让上层回退重扫，
    /// 因为此时本次扫描可能已向 visitor 投出行，回退重扫会导致行双投。
    /// </summary>
    private static string DecodeEntities(string s)
    {
        int first = s.IndexOf('&');
        if (first < 0) return s;

        var sb = new StringBuilder(s.Length);
        sb.Append(s, 0, first);
        int i = first;
        while (i < s.Length)
        {
            char c = s[i];
            if (c != '&') { sb.Append(c); i++; continue; }

            // 引用体最长形态 &#x10FFFF; 共 10 字符；窗口内找不到 ';' 即坏引用
            int semi = -1;
            int maxEnd = Math.Min(s.Length, i + 12);
            for (int j = i + 1; j < maxEnd; j++)
            {
                char d = s[j];
                if (d == ';') { semi = j; break; }
                if (d == '&') break;                        // 嵌套 '&'：前一个不构成引用
            }
            if (semi < 0) throw new System.Xml.XmlException($"未闭合的实体引用：{s}");

            int len = semi - i - 1;
            if (len == 2 && s[i + 1] == 'l' && s[i + 2] == 't') sb.Append('<');
            else if (len == 2 && s[i + 1] == 'g' && s[i + 2] == 't') sb.Append('>');
            else if (len == 3 && s[i + 1] == 'a' && s[i + 2] == 'm' && s[i + 3] == 'p') sb.Append('&');
            else if (len == 4 && s[i + 1] == 'q' && s[i + 2] == 'u' && s[i + 3] == 'o' && s[i + 4] == 't') sb.Append('"');
            else if (len == 4 && s[i + 1] == 'a' && s[i + 2] == 'p' && s[i + 3] == 'o' && s[i + 4] == 's') sb.Append('\'');
            else if (len >= 2 && s[i + 1] == '#')
            {
                if (!TryParseCharRef(s, i + 2, semi, out int value))
                    throw new System.Xml.XmlException($"非法字符引用：{s}");
                if (value > 0x10FFFF
                    || (value < 0x20 && value != 9 && value != 10 && value != 13)
                    || (value >= 0xD800 && value <= 0xDFFF))
                    throw new System.Xml.XmlException($"字符引用越界或非法：{s}");   // 与 XmlReader 对非法字符引用的行为一致
                if (value > 0xFFFF)
                {
                    sb.Append(char.ConvertFromUtf32(value)); // 增补平面：UTF-16 代理对
                    i = semi + 1;
                    continue;
                }
                sb.Append((char)value);
            }
            else throw new System.Xml.XmlException($"未定义的实体：{s}");   // 未定义实体（&nbsp; 等）在无 DTD 的 XmlReader 下同样抛 XmlException

            i = semi + 1;
        }
        return sb.ToString();
    }

    /// <summary>解析字符引用数值：<c>#dddd;</c> 十进制或 <c>#xhhhh;</c> 十六进制（start 指向 '#' 之后）。</summary>
    private static bool TryParseCharRef(string s, int start, int end, out int value)
    {
        value = 0;
        bool hex = s[start] == 'x' || s[start] == 'X';
        int d0 = hex ? start + 1 : start;
        if (d0 >= end || end - d0 > 8) return false;
        for (int i = d0; i < end; i++)
        {
            char c = s[i];
            int d;
            if (hex)
                d = c >= '0' && c <= '9' ? c - '0'
                  : c >= 'a' && c <= 'f' ? c - 'a' + 10
                  : c >= 'A' && c <= 'F' ? c - 'A' + 10 : -1;
            else
                d = c >= '0' && c <= '9' ? c - '0' : -1;
            if (d < 0) return false;
            value = value * (hex ? 16 : 10) + d;
        }
        return true;
    }

    // ───────────────────────── 值解析（与 XlsxSheetReader 共用 XlsxValueParsing） ─────────────────────────

    /// <summary>G15 原样快路径（静态委托，热路径上零分配）；返回 null = 不适用</summary>
    private static readonly Func<string, string?> FastNumericPassthrough =
        static raw => TryFastNumericPassthrough(raw, out string? v) ? v : null;

    /// <summary>
    /// 数字规整快路径：double 只保证 15 位有效十进制数字，因此「规范形式的纯数字且有效位数 ≤ 15」
    /// 做 <c>ToString("G15")</c> 后必然与原样一致 —— 直接原样返回，跳过 parse+format
    /// （实测该段占扫描器内最大单项开销，1490 → 595 ms）。
    /// 含 <c>e/E</c>、千分位、前导零、尾随零、悬空小数点等非规范形式一律走原路径。
    /// </summary>
    private static bool TryFastNumericPassthrough(string raw, out string? value)
    {
        value = null;

        int len = raw.Length;
        if (len == 0) return false;

        int digits = 0;
        bool sawDot = false;
        for (int k = 0; k < len; k++)
        {
            char c = raw[k];
            if (c >= '0' && c <= '9') { digits++; continue; }
            if (c == '.' && !sawDot) { sawDot = true; continue; }
            if (k == 0 && (c == '-' || c == '+')) continue;
            return false;                       // 含 e/E、空白、千分位等 → 原路径
        }

        if (digits == 0 || digits > 15) return false;

        int signLen = (raw[0] == '-' || raw[0] == '+') ? 1 : 0;
        int dot = raw.IndexOf('.');
        int intEnd = dot < 0 ? len : dot;
        int intLen = intEnd - signLen;

        if (intLen <= 0) return false;                                  // "." / "-." → 原路径
        if (intLen > 1 && raw[signLen] == '0') return false;             // "007" → G15 = "7"
        if (intLen == 1 && raw[signLen] == '0' && sawDot)
        {
            value = raw;                                                 // "0.5" 形态：G15 原样
            return true;
        }
        if (sawDot)
        {
            int fracLen = len - dot - 1;
            if (fracLen == 0) return false;                              // "5." → G15 = "5"
            if (raw[len - 1] == '0') return false;                       // "5.10" → G15 = "5.1"
        }

        value = raw;
        return true;
    }
}

/// <summary>
/// 行回调签名。与 <c>XlsxSheetReader.RowVisitor</c> 同形，
/// 单独定义以免为跨类型复用而提升该委托的可见度。
/// </summary>
internal delegate bool RowVisitorLike(int physicalRow, XlsxRowBuffer row);
