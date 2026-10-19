using System.Globalization;
using System.IO;
using System.Text;

namespace ExcelQueryTool.Services
{
    /// <summary>
    /// CSV / TXT 表格文本读取器（自写）。
    ///
    /// 不经 MiniExcel 的理由与 xlsx 侧相同：它的行读取按行建 ExpandoObject
    /// （~11 KB/行分配），且不暴露引号/编码控制。文本解析量级小，
    /// 这里追求行为正确而不是极限吞吐。
    ///
    /// 三项自动识别：
    /// - 分隔符：候选 , \t ; | ，对前 20 行分别试解析，按「出现行占比 × 字段数一致性」打分；
    ///   全部落空时退化为单列表（每行一个字段）。
    /// - 编码：BOM（UTF-8 / UTF-16）→ 无 BOM 时 UTF-8 严格试解（截到最后一个完整换行，
    ///   避免多字节字符在采样边界被切开造成误判）→ 不合法则 GB18030（中文 Windows
    ///   导出 CSV 的常见编码，ASCII 部分与 UTF-8 兼容，故作为兜底安全）。
    /// - 表头：<b>首行必须列名齐全</b>（尾随空列除外，有空列即不是表头）且<b>不可能是纯数字</b>；
    ///   其次看列名是否跨行重复、首两行数字占比是否均 ≥0.5；仍模糊时默认有表头。
    ///
    /// 字段解析遵循 RFC 4180：双引号包裹、引号内逗号/换行（多行字段）、"" 转义。
    ///
    /// 列命名与 xlsx 侧同一套规则（见 <see cref="ExcelDataManager.GetColumnMetadataWithLayout"/>）：
    /// 空表头 → &lt;列N&gt;，重名 → _2/_3。列数 = 全文件最大字段数（保守，与
    /// xlsx「空表头列一律保留」同一硬约束：元数据与数据两条路径列数必须一致）；
    /// 短行补空串，数据行超宽说明文件在两次扫描之间被改写 → 抛异常（与 xlsx 列数守卫同语义）。
    /// </summary>
    public static class TextTableReader
    {
        private static readonly char[] DelimiterCandidates = [',', '\t', ';', '|'];

        static TextTableReader()
        {
            // GB18030 等 Windows 代码页在 .NET 上需注册 CodePages 提供程序
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public static bool IsTextFile(string filePath)
        {
            string ext = Path.GetExtension(filePath);
            return ext.Equals(".csv", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".txt", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 表结构。列数覆盖全文件所有行（含表头行）的最大字段数 —— 与数据路径
        /// <see cref="EnumerateRows"/> 的定宽投影共用同一数值，这是硬约束。
        /// Encoding / Delimiter 一并带回，供第二趟解析复用，免于重复探测。
        /// </summary>
        public sealed class TextLayout
        {
            public required int ColumnCount { get; init; }
            public required bool HasHeader { get; init; }

            /// <summary>表头行原文字段（已 Trim）；HasHeader=false 时为空数组</summary>
            public string?[] Headers { get; init; } = [];

            /// <summary>第二趟解析复用的编码与分隔符（ReadLayout 探测结果，免重复探测）</summary>
            public required Encoding Encoding { get; init; }
            public required char Delimiter { get; init; }

            /// <summary>源文件路径。供 <see cref="BuildColumns"/> 做数字列抽样（需再次读文件）。</summary>
            public string FilePath { get; init; } = "";

            /// <summary>
            /// 字节级快路径数据：仅 <see cref="TryReadFast"/> 成功时非 null，char 级
            /// <see cref="ReadLayout"/> 产出恒为 null。数据段据此按块解码，免去第二遍读盘+扫描。
            /// </summary>
            public TextFastData? Fast { get; init; }
        }

        /// <summary>表头判定采样行数上限（重复率统计用；再多行收益递减）</summary>
        private const int HeaderSampleRows = 8;

        /// <summary>
        /// 全量扫描一次文件得到表结构。文本文件解析成本低，两趟扫描
        /// （本方法 + <see cref="EnumerateRows"/>）换取「列数在产出前定案」的简单性。
        /// </summary>
        public static TextLayout ReadLayout(string filePath)
        {
            Encoding encoding = DetectEncoding(filePath);
            char delimiter = DetectDelimiter(filePath, encoding);

            var sample = new List<string?[]>(HeaderSampleRows);
            int maxCols = 0;

            using (var reader = new StreamReader(filePath, encoding))
            {
                foreach (var fields in ParseRows(reader, delimiter, null))
                {
                    if (fields.Count > maxCols) maxCols = fields.Count;
                    if (sample.Count < HeaderSampleRows) sample.Add([.. fields]);
                }
            }

            bool hasHeader = DetectHeader(sample);
            string?[]? firstRow = sample.Count > 0 ? sample[0] : null;
            return new TextLayout
            {
                ColumnCount = Math.Max(maxCols, 1),
                HasHeader = hasHeader,
                Headers = hasHeader && firstRow != null ? firstRow : [],
                Encoding = encoding,
                Delimiter = delimiter,
                FilePath = filePath,
            };
        }

        /// <summary>
        /// 第二趟流式产出定宽行（长度恒为 <see cref="TextLayout.ColumnCount"/>，短行补空串）。
        /// HasHeader 时跳过首行（首行已被 <see cref="ReadLayout"/> 用作表头）。
        /// 行超宽 = 文件在两次扫描之间被改写 → 抛异常，不按过时映射静默错位。
        /// </summary>
        public static IEnumerable<string?[]> EnumerateRows(string filePath, TextLayout layout, CancellationToken token)
        {
            bool headerPending = layout.HasHeader;
            int rowNo = 0;

            using var reader = new StreamReader(filePath, layout.Encoding);
            foreach (var fields in ParseRows(reader, layout.Delimiter, null))
            {
                token.ThrowIfCancellationRequested();
                rowNo++;

                if (headerPending)
                {
                    headerPending = false;
                    continue;
                }

                if (fields.Count > layout.ColumnCount)
                    throw new InvalidOperationException(
                        $"文本文件列结构已变化（第 {rowNo} 行出现 {fields.Count} 列，表头 {layout.ColumnCount} 列），请重新打开文件");

                var row = new string?[layout.ColumnCount];
                for (int i = 0; i < fields.Count; i++) row[i] = fields[i];
                yield return row;
            }
        }

        // ---------- 字节级快路径（UTF-8 / GB18030） ----------
        //
        // 核心是 <see cref="ParseRowsRanges"/> —— <see cref="ParseRows"/>（字符级状态机）在字节上的
        // 直接移植：同一套「引号感知 / "" 转义 / CRLF 合并 / 空行跳过 / EOF 收尾」判定，只是
        // 在字节上还原字段值而非复制字符。故行集合、每行字段数、字段值均与字符路径逐字一致。
        // 安全前提：字段/行边界只认「引号、分隔符、换行」这几类 ASCII 单字节，且它们绝不能
        // 出现在某个多字节字符的内部。
        //   · UTF-8：多字节序列的每个字节都 ≥ 0x80，故任何 ASCII 边界字节都不会落在字符中间 —— 恒安全。
        //   · GB18030：2 字节尾 0x40-0x7E∪0x80-0xFE、4 字节 2/4 位 0x30-0x39、3 位 0x81-0xFE。
        //     分隔符 , (0x2C) / ; (0x3B) / \t (0x09) 均不落在上述任何集合 → 恒安全；
        //     但 | (0x7C) 落在 2 字节尾 0x40-0x7E 内，理论上可出现在 GB18030 字符中间 → 不安全。
        // 故 GB18030 仅在探测出的分隔符 ≠ | 时走字节路径，其余（UTF-16、| 分隔的 GB18030、
        // 其它代码页）一律回退 char 级。该门槛同时保证分隔符探测本身无伪分裂（见
        // <see cref="DetectDelimiterBytes"/> 与 <see cref="TryReadFast"/> 的说明）。

        /// <summary>
        /// 字节级快路径数据。由 <see cref="TryReadFast"/> 产出，持有整文件字节与全部被产出行的
        /// 字节区间（绝对偏移，已含 BOM 跳过），数据段据此按块并行解码。
        /// </summary>
        public sealed class TextFastData
        {
            /// <summary>整文件字节</summary>
            public required byte[] Bytes { get; init; }

            /// <summary>被产出行的字节区间 [s, e)（绝对偏移；e 为行内容截止处，已剔除 \r\n / 孤立 \r）</summary>
            public required (int s, int e)[] RowRanges { get; init; }

            /// <summary>编码（与 char 路径同一实例语义）</summary>
            public required Encoding Encoding { get; init; }

            /// <summary>分隔符</summary>
            public required char Delimiter { get; init; }

            /// <summary>
            /// 把全局行下标区间 [startRow, endRow) 的数据行批量解码为定宽字段数组列表（长度 =
            /// <paramref name="columnCount"/>，短行以 null 补齐）—— 与 <see cref="EnumerateRows"/> 的单行
            /// 产出完全同形（引号/转义/换行语义逐字一致）。供并行 worker 调用。
            ///
            /// 相比「逐行解码」，本方法把整个字节段<b>一次性解码</b>
            /// （每块仅一次 encoding 解码，而非逐行/逐字段的密集小解码），再用
            /// <see cref="ParseRows"/>（char 路径<b>同一</b>状态机）切分 —— 既消除逐行分配淹没并行，
            /// 又因复用同一状态机而零移植偏差。块边界必落在行边界上（<see cref="RowRanges"/> 每段为一整行），
            /// 且 UTF-8/GB18030 的行边界即字符边界（见文件头安全性说明），故整块解码无多字节截断。
            /// </summary>
            public List<string?[]> DecodeRowsChunk(int startRow, int endRow, int columnCount, CancellationToken token)
            {
                var result = new List<string?[]>(Math.Max(0, endRow - startRow));
                if (endRow <= startRow) return result;

                // 块字节段：首行起点 → 末行内容终点（绝对偏移，已含 BOM 跳过）
                int s = RowRanges[startRow].s;
                int e = RowRanges[endRow - 1].e;

                // 整块一次性解码（块边界即字符边界，见文件头说明）
                string text = Encoding.GetString(Bytes, s, e - s);
                char del = Delimiter;
                int rowInChunk = 0;

                using var reader = new StringReader(text);
                foreach (var fields in ParseRows(reader, del, null))
                {
                    token.ThrowIfCancellationRequested();

                    // 块内字段数与布局阶段逐字节判定一致；超宽 = 文件在两次读取间被改写
                    if (fields.Count > columnCount)
                        throw new InvalidOperationException(
                            $"文本文件列结构已变化（第 {startRow + rowInChunk + 1} 行出现 {fields.Count} 列，表头 {columnCount} 列），请重新打开文件");

                    var row = new string?[columnCount];
                    for (int c = 0; c < fields.Count; c++) row[c] = fields[c];
                    result.Add(row);
                    rowInChunk++;
                }

                return result;
            }
        }

        /// <summary>
        /// 字节级读取表结构：一次字节扫描（<see cref="ParseRowsRanges"/>，ParseRows 的字节移植）
        /// 得到列数与表头，免去 char 级全量扫描。仅当 <see cref="IsByteLevelEncoding"/> 成立时返回
        /// true（否则回退 <see cref="ReadLayout"/>）。列数口径与 char 路径 <c>Math.Max(全文件最大字段数, 1)</c> 完全一致。
        /// </summary>
        public static bool TryReadFast(string filePath, out TextLayout layout)
        {
            layout = null!;
            Encoding encoding = DetectEncoding(filePath);
            if (!IsByteLevelEncoding(encoding)) return false;

            byte[] data = File.ReadAllBytes(filePath);
            int offset = 0;
            if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF)
                offset = 3;   // UTF-8 BOM 非字段内容，与 StreamReader 行为一致地跳过

            char delimiter = DetectDelimiterBytes(data, offset);

            // GB18030 + | 分隔：| (0x7C) 可出现在 GB18030 2 字节字符的尾部，字节边界会误判 → 回退 char 级。
            // （UTF-8 不受此限：多字节字节恒 ≥ 0x80。）
            if (delimiter == '|' && !encoding.WebName.Equals("utf-8", StringComparison.OrdinalIgnoreCase))
                return false;

            var ranges = new List<int>();
            var fieldCounts = new List<int>();
            ParseRowsRanges(data, offset, delimiter, 0, ranges, fieldCounts, null, encoding);
            if (fieldCounts.Count == 0) return false;   // 空文件 → char 路径

            int maxCols = fieldCounts.Max();            // 与 ReadLayout：全文件最大字段数
            bool hasHeader = DetectHeader(DecodeHeaderSample(data, ranges, delimiter, encoding));
            string?[]? firstRow = hasHeader ? DecodeHeaderRow(data, ranges[0], ranges[1], delimiter, encoding) : null;

            var fast = new TextFastData
            {
                Bytes = data,
                RowRanges = ToPairs(ranges),
                Encoding = encoding,
                Delimiter = delimiter,
            };
            layout = new TextLayout
            {
                ColumnCount = Math.Max(maxCols, 1),
                HasHeader = hasHeader,
                Headers = firstRow ?? [],
                Encoding = encoding,
                Delimiter = delimiter,
                FilePath = filePath,
                Fast = fast,
            };
            return true;
        }

        /// <summary>交错区间数组 [s0,e0,s1,e1,…] → (s,e)[]。</summary>
        private static (int s, int e)[] ToPairs(List<int> r)
        {
            var pairs = new (int s, int e)[r.Count / 2];
            for (int i = 0; i < pairs.Length; i++) pairs[i] = (r[2 * i], r[2 * i + 1]);
            return pairs;
        }

        /// <summary>表头判定采样（最多 <see cref="HeaderSampleRows"/> 行），逐行解码，口径同 <see cref="ReadLayout"/> 的 sample。</summary>
        private static List<string?[]> DecodeHeaderSample(byte[] data, List<int> rowRanges, char del, Encoding encoding)
        {
            var sample = new List<string?[]>(HeaderSampleRows);
            int n = Math.Min(HeaderSampleRows, rowRanges.Count / 2);
            for (int i = 0; i < n; i++)
                sample.Add(DecodeHeaderRow(data, rowRanges[2 * i], rowRanges[2 * i + 1], del, encoding));
            return sample;
        }

        private static string?[] DecodeHeaderRow(byte[] data, int s, int e, char del, Encoding encoding)
        {
            var fields = new List<string>();
            ParseRowsRanges(data, s, del, 0, null, null, fields, encoding, e);
            return [.. fields];
        }

        /// <summary>
        /// <see cref="ParseRows"/> 的字节级直接移植。逐字节推进状态机，对每个被产出行记录
        /// 字节区间（<paramref name="rowRanges"/>）与字段数（<paramref name="fieldCounts"/>），
        /// 可选地把字段值解码到 <paramref name="outFields"/>（按行解码字段值用，见 <see cref="DecodeHeaderRow"/>）。
        /// 判定与 <see cref="ParseRows"/> 逐字一致：
        /// <list type="bullet">
        ///   <item>字段起始引号 / <c>""</c> 转义 / 中间引号按字面 / 引号内分隔符与换行皆为字段内容；</item>
        ///   <item>引号外 <paramref name="del"/> 收尾一个字段；</item>
        ///   <item><c>\r\n</c> 合并为一个行终结符；<c>\r</c> / <c>\n</c> / EOF 均收尾当前行；</item>
        ///   <item>「单字段且为空」的行（尾随换行 / 空行分隔）跳过，与字符版同口径；</item>
        ///   <item>EOF 收尾：有未终结字段则补收一行（字段数 = 分隔符数 + 1，与字符版一致）。</item>
        /// </list>
        /// <paramref name="maxRows"/> 为产出行数上限（0 = 不限），仅对「计数/收集」生效。
        /// </summary>
        private static void ParseRowsRanges(
            byte[] data, int offset, char del, int maxRows,
            List<int>? rowRanges, List<int>? fieldCounts, List<string>? outFields,
            Encoding encoding, int end = -1)
        {
            int n = end < 0 ? data.Length : end;
            int db = (int)del;
            int inQuotes = 0;
            int emitted = 0;
            int rowStart = offset;
            int fieldCount = 0;      // 当前行已收尾字段数（对应 ParseRows 的 fields.Count，收尾后含末字段）

            // current 字段值的字节缓冲（对应 ParseRows 的 current StringBuilder）。
            // 引号语义在此逐字节还原：起始引号不入户、"" 转义→单个 "、中间引号按字面 ——
            // 只有状态机能产出正确字段值（引号字段值不是连续子串，不能靠区间截取）。
            var cur = new byte[1024];
            int curLen = 0;

            void AppendByte(int b)
            {
                if (outFields != null)
                {
                    if (curLen == cur.Length) Array.Resize(ref cur, cur.Length * 2);
                    cur[curLen] = (byte)b;
                }
                curLen++;
            }

            // 收尾当前字段：解码入 outFields（value 模式）+ 计数；返回该字段长度（空行判定用）。
            int EndField()
            {
                int len = curLen;
                if (outFields != null) outFields.Add(len > 0 ? encoding.GetString(cur, 0, len) : "");
                fieldCount++;
                curLen = 0;
                return len;
            }

            for (int i = offset; i < n; i++)
            {
                int c = data[i];
                if (inQuotes != 0)
                {
                    if (c == '"')
                    {
                        int next = i + 1 < n ? data[i + 1] : -1;
                        if (next == '"') { AppendByte('"'); i++; }   // "" 转义 → 字面引号
                        else inQuotes = 0;
                    }
                    else AppendByte(c);      // 引号内的分隔符/换行/内容皆为字段值
                    continue;
                }
                if (c == '"')
                {
                    if (curLen == 0) inQuotes = 1;   // 字段起始引号（非字段内容）
                    else AppendByte(c);              // 字段中间引号按字面保留
                    continue;
                }
                if (c == db)
                {
                    EndField();
                    continue;
                }
                if (c == '\r' || c == '\n')
                {
                    int eContent = i;                    // 行内容截止于行终结符之前（\r / \n 不入字段）
                    if (c == '\r' && i + 1 < n && data[i + 1] == '\n') i++;   // \r\n 合并为一个行终结符
                    int len = EndField();

                    if (outFields == null)               // layout 模式：空行跳过，否则记录行
                    {
                        if (!(fieldCount == 1 && len == 0))
                        {
                            emitted++;
                            rowRanges?.Add(rowStart); rowRanges?.Add(eContent);
                            fieldCounts?.Add(fieldCount);
                        }
                        fieldCount = 0;
                        rowStart = i + 1;                // 空行也要推进行起点，否则下一行区间会含被跳过的换行
                        if (maxRows != 0 && emitted >= maxRows) break;
                    }
                    continue;
                }
                AppendByte(c);                          // 其余字节为字段内容
            }

            // EOF：收尾未终结的字段/行（与 ParseRows 的 EOF 块逐字一致）
            if (curLen > 0 || fieldCount > 0)
            {
                int len = EndField();
                if (outFields == null && !(fieldCount == 1 && len == 0))
                {
                    rowRanges?.Add(rowStart); rowRanges?.Add(n);
                    fieldCounts?.Add(fieldCount);
                }
            }
        }

        /// <summary>
        /// 分隔符探测（字节级）：候选分隔符按「多列行占比 ×2 + 一致性」打分，取最高分，
        /// 与 <see cref="DetectDelimiter"/>（char 级）同口径 —— 复用 <see cref="ParseRowsRanges"/> 取前 20 个非空行的字段数。
        /// </summary>
        private static char DetectDelimiterBytes(byte[] data, int offset)
        {
            char best = ',';
            double bestScore = -1;

            foreach (char candidate in DelimiterCandidates)
            {
                var counts = new List<int>();
                ParseRowsRanges(data, offset, candidate, 20, null, counts, null, Encoding.ASCII);

                int multiColLines = counts.Count(c => c >= 2);
                if (multiColLines == 0) continue;

                double mean = counts.Average();
                double variance = counts.Sum(c => (c - mean) * (c - mean)) / counts.Count;
                double consistency = mean > 0 ? Math.Max(0, 1 - Math.Sqrt(variance) / mean) : 0;
                double coverage = (double)multiColLines / counts.Count;

                double score = coverage * 2 + consistency;
                if (score > bestScore) { bestScore = score; best = candidate; }
            }
            return best;
        }

        /// <summary>
        /// 编码是否为「字节级候选」：UTF-8（恒安全）或 GB18030（分隔符 ≠ <c>|</c> 时安全，
        /// 见文件头说明与 <see cref="TryReadFast"/> 的最终判定）。UTF-16（含 BOM-less）与其余
        /// 代码页逐字符非 1 字节、无法按字节定位边界，恒回退 char 级。
        /// </summary>
        public static bool IsByteLevelEncoding(Encoding encoding)
        {
            string name = encoding.WebName ?? "";
            return name.Equals("utf-8", StringComparison.OrdinalIgnoreCase)
                || name.Equals("gb18030", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 列元数据。命名规则与 <see cref="ExcelDataManager.GetColumnMetadataWithLayout"/> 一致：
        /// 空表头 → &lt;列N&gt;，重名追加 _2/_3（首次出现不加后缀）。
        /// 文本无图片列（IsImage=false），列宽用程序默认（ExcelWidth=0）。
        /// 数字列判定与 xlsx 侧同口径：抽样若干数据行，非空值中数字占比 ≥ 0.8 即为数字列。
        /// </summary>
        public static List<ColumnMeta> BuildColumns(TextLayout layout)
        {
            var columns = new List<ColumnMeta>(layout.ColumnCount);
            var normalizer = new HeaderNormalizer();
            var numeric = DetectNumericColumns(layout);

            for (int col = 1; col <= layout.ColumnCount; col++)
            {
                string header = layout.HasHeader && col <= layout.Headers.Length
                    ? layout.Headers[col - 1]?.Trim() ?? ""
                    : "";

                string name = normalizer.Normalize(col, header);

                columns.Add(new ColumnMeta(name, isImage: false, excelWidth: 0, isNumeric: numeric[col]));
            }

            return columns;
        }

        /// <summary>
        /// 列级数字标记（1 基，[0] 未使用）。抽样前 <see cref="NumericColumnClassifier.SampleRows"/> 行数据，
        /// 非空值中数字占比 ≥ 0.8 判为数字列；样本不足 3 个非空值时不判定（保持文本左对齐）。
        /// 口径统一在 <see cref="NumericColumnClassifier"/>，与 xlsx 路径共用同一实现。
        /// </summary>
        private static bool[] DetectNumericColumns(TextLayout layout)
        {
            var result = new bool[layout.ColumnCount + 1];
            var total = new int[layout.ColumnCount + 1];
            var hits = new int[layout.ColumnCount + 1];

            int rows = 0;
            foreach (var fields in EnumerateRows(layout.FilePath, layout, CancellationToken.None))
            {
                if (++rows > NumericColumnClassifier.SampleRows) break;

                for (int c = 0; c < layout.ColumnCount && c < fields.Length; c++)
                {
                    string? v = fields[c];
                    if (string.IsNullOrWhiteSpace(v)) continue;
                    total[c + 1]++;
                    if (IsNumericText(v.Trim())) hits[c + 1]++;
                }
            }

            for (int c = 1; c <= layout.ColumnCount; c++)
                result[c] = NumericColumnClassifier.Classify(hits[c], total[c]);

            return result;
        }

        /// <summary>文本是否为数字型（列级对齐判定用）。只认纯数字/千分位/正负号/百分号/小数点。</summary>
        private static bool IsNumericText(string text)
            => text.Length > 0 && double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out _);

        // ---------- 表头判定 ----------

        /// <summary>
        /// 表头判定，按以下优先级：
        ///
        /// <b>1. 首行存在空列 = 不是表头（硬前提）</b>
        /// 表头的语义是「每列都有列名」。首行只要出现空列，就不满足表头条件，
        /// 判「无表头」。这条规则最直观、可解释，实测 `主数据导入.txt` 正是如此：
        /// 74 列中首行 43 列为空 → 直接判无表头。
        ///
        /// 唯一的例外是<b>尾随空列</b>：CSV/TXT 常在行尾留下多余的分隔符
        /// （`a,b,c,,`），这种「右侧连续空列」是格式残留而非缺少列名，
        /// 不参与判定；其余位置（即尾随空列之前）若出现空值则视为不是表头。
        ///
        /// <b>2. 首行全为数字 = 不是表头</b>
        /// 列名总归是文字，不可能整行都是数字（`1,2,3` 是数据不是列名）。
        /// 注意条件是「<b>全部</b>字段皆为数字」，而非「数字过半」——
        /// 真实的列名常带数字（`数量1`、`2024年`、`金额(元)`），
        /// 只要有一个字段不是数字就不能据此否定表头。
        ///
        /// <b>3. 列名跨行重复 = 无表头</b>
        /// 代码型数据行（`ROH`、`PC`、`X` 等代码列）会在后续行反复出现同列同值，
        /// 而列名只出现在表头那一行。用于拦住「首行恰好填满但实为数据」的情形。
        ///
        /// <b>4. 数字占比兜底（保守）</b>
        /// 仅当首行<b>过半字段是数字且无任何文字样式的列名</b>时，
        /// 才结合次行占比判「无表头」。
        ///
        /// 这条判据只在正面判据（1、2、3）均不成立时才走到，因此必须保守：
        /// 曾因「首行与次行占比均 ≥0.5」过宽，把 `名称,100,200`（真实列名 + 数字列）
        /// 与 `2024年,2025年`（年份列名）误判为无表头。改为要求
        /// 「首行数字字段数 ≥ 有值字段数的一半 <b>且首行存在非数字字段时该字段不像列名</b>」
        /// 后不再误伤。仍模糊时默认<b>有表头</b> —— 数据导出类文件几乎都带表头，
        /// 误判为「无表头」会把首行数据吃成列名，代价更高。
        /// </summary>
        private static bool DetectHeader(List<string?[]> sample)
        {
            if (sample.Count == 0) return false;   // 空文件 → 无表头（空表）

            var first = sample[0];

            // 判据 1：表头要求列名齐全（尾随空列除外）
            if (HasEmptyBeforeTrailing(first)) return false;

            // 判据 2：列名不可能是纯数字（整行皆为数字 = 数据行）
            if (IsAllNumeric(first)) return false;

            // 判据 3：列名不会跨行重复出现
            if (sample.Count > 1 && RepeatRatio(first, sample) >= 0.6) return false;

            // 判据 4：数字占比兜底 —— 要求「绝大多数字段是数字」且首行无类列名文字，
            // 避免把带数字的真实列名（`数量1`、`2024年`）误判成数据
            if (sample.Count > 1 && IsNumericDominant(first) && IsNumericDominant(sample[1])) return false;

            return true;
        }

        /// <summary>
        /// 首行是否「以数字为绝对多数」（≥80%）且不含类列名的文字字段。
        ///
        /// 阈值取 0.8 而非 0.5：`名称,100,200` 这类真实表头（列名 + 数字列混合）
        /// 占比 0.67，不能被判成数据；只有 `1,2,3`、`100,200,300` 这种
        /// 几乎全数字的行才符合「数据行」特征。纯数字行已被判据 2 拦下，
        /// 本判据实际兜的是「含少量文字但仍是数据」的行（如 `A,1,2,3,4`）。
        /// </summary>
        private static bool IsNumericDominant(string?[] fields)
        {
            int nonEmpty = 0, numeric = 0;
            foreach (var f in fields)
            {
                string t = f?.Trim() ?? "";
                if (t.Length == 0) continue;

                nonEmpty++;
                if (IsNumericLike(t)) numeric++;
            }

            if (nonEmpty == 0) return false;
            return (double)numeric / nonEmpty >= 0.8;
        }

        /// <summary>
        /// 首行是否整行皆为数字/日期（<b>所有</b>有值字段）。
        ///
        /// 列名是文字，`1,2,3` 这样的行只能是数据。要求「全部」而非「过半」：
        /// 列名带数字是常见现象（`数量1`、`2024年`、`金额(元)`），
        /// 有一个非数字字段就不能据此认定不是表头。
        /// 全空行（尾随空列之外无任何值）不在此列，返回 false 交由后续判据处理。
        /// </summary>
        private static bool IsAllNumeric(string?[] fields)
        {
            int nonEmpty = 0;
            foreach (var f in fields)
            {
                string t = f?.Trim() ?? "";
                if (t.Length == 0) continue;

                nonEmpty++;
                if (!IsNumericLike(t)) return false;
            }

            return nonEmpty > 0;
        }

        /// <summary>
        /// 首行在「尾部连续空列之前」是否存在空值。
        ///
        /// 表头要求列名齐全。尾随空列（行尾多余分隔符造成的右侧连续空值）属
        /// 格式残留，予以忽略；此外任何位置的空列都说明首行不是表头。
        /// 数据行常见这种情况：`主数据导入.txt` 首行 31 列有值、43 列空，
        /// 空值分布在尾随段之前（如第 4、8、9 列），故判为非表头。
        /// </summary>
        private static bool HasEmptyBeforeTrailing(string?[] first)
        {
            // 找到最后一个有值的位置，其之前（含）非空即为「非尾随空列」
            int lastPopulated = -1;
            for (int c = first.Length - 1; c >= 0; c--)
            {
                if ((first[c]?.Trim() ?? "").Length > 0) { lastPopulated = c; break; }
            }

            for (int c = 0; c <= lastPopulated; c++)
            {
                if ((first[c]?.Trim() ?? "").Length == 0) return true;
            }

            return false;
        }

        /// <summary>
        /// 首行「文字字段」中，在后续样本行里出现同列相同取值的比例。
        ///
        /// 只在首行字段本身是文字（非数字/日期）时才计入分母 —— 数字列（数量、编号）
        /// 跨行相同可能是巧合，不构成「这一行是数据」的证据；而代码列（`ROH`、`PC`、`X`）
        /// 跨行相同才是数据行的典型特征。
        /// 单个字段以「后续行中至少一半与其相同」为记为重复（6 行样本需 ≥3 行相同，
        /// 但首行本身不计入，故按后续行数计），避免个别行差异拉低整体。
        /// 首行无文字字段时返回 0（无从判断，交给数字占比兜底）。
        /// </summary>
        private static double RepeatRatio(string?[] first, List<string?[]> sample)
        {
            int following = sample.Count - 1;
            if (following <= 0) return 0;

            int threshold = Math.Max(1, (following + 1) / 2);
            int textFields = 0, repeated = 0;

            for (int c = 0; c < first.Length; c++)
            {
                string a = first[c]?.Trim() ?? "";
                if (a.Length == 0 || IsNumericLike(a)) continue;

                textFields++;

                int same = 0;
                for (int r = 1; r < sample.Count; r++)
                {
                    if (c < sample[r].Length && (sample[r][c]?.Trim() ?? "") == a) same++;
                }
                if (same >= threshold) repeated++;
            }

            return textFields == 0 ? 0 : (double)repeated / textFields;
        }

        /// <summary>
        /// 字段是否「数字样」：纯数字或标准日期格式。
        ///
        /// 刻意不用宽松的 <c>DateTime.TryParse</c> —— 它会把 <c>2024年</c>、<c>1月</c>
        /// 这类<b>带中文后缀的年份/月份</b>解析成日期，而它们通常是列名而非日期值
        /// （实测 `2024年,2025年` 曾被误判为「全数字行」→ 错判无表头）。
        /// 因此先试数字，再只认「纯数字构成的日期串」（如 `2024-01-01`、`2024/1/1`）。
        /// </summary>
        private static bool IsNumericLike(string text)
        {
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out _)) return true;

            // 日期只认「数字 + 常见分隔符」构成的形式，拒绝中文/字母后缀
            foreach (char c in text)
            {
                if (!char.IsDigit(c) && c is not ('-' or '/' or '.' or ':' or ' ')) return false;
            }

            return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
        }

        // ---------- 分隔符嗅探 ----------

        /// <summary>
        /// 对每个候选分隔符取前 20 行试解析：出现多列的行占比（权重 2）+
        /// 字段数一致性（行间字段数越接近得分越高）。全部落空时返回逗号
        /// （解析结果为单列表，行为安全）。
        /// </summary>
        private static char DetectDelimiter(string filePath, Encoding encoding)
        {
            char best = ',';
            double bestScore = -1;

            foreach (char candidate in DelimiterCandidates)
            {
                var counts = new List<int>();
                using (var reader = new StreamReader(filePath, encoding))
                {
                    foreach (var fields in ParseRows(reader, candidate, maxRows: 20))
                        counts.Add(fields.Count);
                }

                int multiColLines = counts.Count(c => c >= 2);
                if (multiColLines == 0) continue;

                double mean = counts.Average();
                double variance = counts.Sum(c => (c - mean) * (c - mean)) / counts.Count;
                double consistency = mean > 0 ? Math.Max(0, 1 - Math.Sqrt(variance) / mean) : 0;
                double coverage = (double)multiColLines / counts.Count;

                double score = coverage * 2 + consistency;
                if (score > bestScore)
                {
                    bestScore = score;
                    best = candidate;
                }
            }

            return best;
        }

        // ---------- 编码检测 ----------

        /// <summary>
        /// BOM 优先；无 BOM 时 UTF-8 严格试解（采样截到最后一个完整换行），
        /// 不合法 → GB18030（GBK 系超集，ASCII 兼容，兜底安全）。
        /// </summary>
        private static Encoding DetectEncoding(string filePath)
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                Span<byte> bom = stackalloc byte[3];
                int n = fs.Read(bom);
                if (n >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
                    return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
                if (n >= 2 && bom[0] == 0xFF && bom[1] == 0xFE) return Encoding.Unicode;
                if (n >= 2 && bom[0] == 0xFE && bom[1] == 0xFF) return Encoding.BigEndianUnicode;
            }

            byte[] prefix = ReadPrefix(filePath, 1024 * 1024);
            int cut = Array.LastIndexOf(prefix, (byte)'\n');
            if (cut > 0) prefix = prefix[..(cut + 1)];

            try
            {
                _ = new UTF8Encoding(false, throwOnInvalidBytes: true).GetString(prefix);
                return new UTF8Encoding(false);
            }
            catch (DecoderFallbackException)
            {
                return Encoding.GetEncoding("GB18030");
            }
        }

        private static byte[] ReadPrefix(string filePath, int maxBytes)
        {
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var buffer = new byte[(int)Math.Min(maxBytes, fs.Length)];
            int read = 0;
            while (read < buffer.Length)
            {
                int n = fs.Read(buffer, read, buffer.Length - read);
                if (n == 0) break;
                read += n;
            }
            return read == buffer.Length ? buffer : buffer[..read];
        }

        // ---------- RFC 4180 状态机 ----------

        /// <summary>
        /// 逐行解析为字段列表。字符级状态机（行式读取会切坏引号内的多行字段）。
        /// 「单字段且为空」的空行跳过（尾部换行/空行分隔）。
        /// 引号未闭合按字面收尾，不抛异常 —— 容错优先于严格性。
        /// </summary>
        private static IEnumerable<List<string>> ParseRows(TextReader reader, char delimiter, int? maxRows)
        {
            var fields = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false;
            int emittedRows = 0;

            while (true)
            {
                int c = reader.Read();
                if (c == -1) break;

                if (inQuotes)
                {
                    if (c == '"')
                    {
                        int next = reader.Peek();
                        if (next == '"')
                        {
                            current.Append('"');   // "" 转义
                            reader.Read();
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        current.Append((char)c);   // 引号内的分隔符/换行都是字段内容
                    }
                    continue;
                }

                if (c == '"')
                {
                    if (current.Length == 0) inQuotes = true;   // 字段起始引号
                    else current.Append('"');                    // 字段中间引号按字面
                }
                else if (c == delimiter)
                {
                    fields.Add(current.ToString());
                    current.Clear();
                }
                else if (c == '\r' || c == '\n')
                {
                    if (c == '\r' && reader.Peek() == '\n') reader.Read();
                    fields.Add(current.ToString());
                    current.Clear();

                    if (!(fields.Count == 1 && fields[0].Length == 0))
                    {
                        emittedRows++;
                        yield return fields;
                    }
                    fields = [];

                    if (maxRows.HasValue && emittedRows >= maxRows.Value) yield break;
                }
                else
                {
                    current.Append((char)c);
                }
            }

            // EOF：收尾未终结的字段/行
            if (current.Length > 0 || fields.Count > 0)
            {
                fields.Add(current.ToString());
                if (!(fields.Count == 1 && fields[0].Length == 0))
                    yield return fields;
            }
        }
    }
}
