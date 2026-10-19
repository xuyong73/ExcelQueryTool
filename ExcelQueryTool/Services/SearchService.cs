using System.Globalization;
using System.Text;

namespace ExcelQueryTool.Services
{
    /// <summary>
    /// 搜索语法解析与匹配。
    ///
    /// 语法：<c>term</c> 整行包含；<c>!term</c> 排除；<c>a&amp;b</c>（或 <c>a+b</c>）并且；
    /// <c>a,b</c>（或 <c>a;b</c>）或者；<c>()</c> 分组；<c>列=值</c> 列限定包含
    /// （列可写表头名或 1 基列序号，如 <c>部门=技术</c>、<c>2=技术</c>）；
    /// <c>列==值</c> 整格等于；<c>列&lt;&gt;值</c>（或 <c>列!=值</c>）不包含；
    /// <c>列&gt;值</c> 等数值比较。中英标点等价（，；！（）＆＋＝）。
    ///
    /// 性能约定：<see cref="Prepare"/> 只调一次，产出表达式树 + 去重词表；
    /// 逐行匹配走两阶段——一趟扫描行内全部单元格填充词命中表（多词不重复扫行），
    /// 再查表求值树。列限定词在 <see cref="BindColumns"/> 绑定显示列下标后按列求值。
    /// </summary>
    public class SearchService
    {
        public abstract class SearchExpr
        {
            /// <summary>rowData：行数据（[0]=行号）；termHits：整行词命中表（Terms 下标）</summary>
            public abstract bool Evaluate(object[] rowData, ReadOnlySpan<bool> termHits);
        }

        public class AndExpr : SearchExpr
        {
            public SearchExpr Left { get; set; } = null!;
            public SearchExpr Right { get; set; } = null!;
            public override bool Evaluate(object[] rowData, ReadOnlySpan<bool> termHits)
                => Left.Evaluate(rowData, termHits) && Right.Evaluate(rowData, termHits);
        }

        public class OrExpr : SearchExpr
        {
            public SearchExpr Left { get; set; } = null!;
            public SearchExpr Right { get; set; } = null!;
            public override bool Evaluate(object[] rowData, ReadOnlySpan<bool> termHits)
                => Left.Evaluate(rowData, termHits) || Right.Evaluate(rowData, termHits);
        }

        public class NotExpr : SearchExpr
        {
            public SearchExpr Child { get; set; } = null!;
            public override bool Evaluate(object[] rowData, ReadOnlySpan<bool> termHits)
                => !Child.Evaluate(rowData, termHits);
        }

        public class TermExpr : SearchExpr
        {
            public string Term { get; set; } = "";
            /// <summary>命中表下标（<see cref="Prepare"/> 时分配）；-1 = 空词（匹配一切）</summary>
            public int Index { get; set; } = -1;
            /// <summary>空词（Index=-1）匹配一切；否则取该词在本行的命中位</summary>
            public override bool Evaluate(object[] rowData, ReadOnlySpan<bool> termHits)
                => Index < 0 ? true : termHits[Index];
        }

        /// <summary>
        /// 列限定词：列&lt;op&gt;值。op：<c>=</c> 包含（默认）、<c>==</c> 整格等于（两侧 Trim 后忽略大小写）、
        /// <c>&lt;&gt;</c> 不包含（等价 <c>!列=值</c>）、
        /// <c>&gt;</c> <c>&lt;</c> <c>&gt;=</c> <c>&lt;=</c> 数值比较（任一侧非数值则该行不匹配，日期按 yyyy-MM-dd 文本不可比）。
        /// Index 未绑定或越界时永不匹配（列名写错 / 列序号超界可预期地查无结果，BindColumns 会登记进 UnresolvedColumns）。
        /// </summary>
        public class ColumnTermExpr : SearchExpr
        {
            public string Column { get; set; } = "";
            public string Value { get; set; } = "";
            /// <summary>""=包含；"=="等于；"&lt;&gt;"不包含；"&gt;" "&lt;" "&gt;=" "&lt;=" 数值比较</summary>
            public string Op { get; set; } = "";
            /// <summary>运算符为 &lt;&gt; 时值为排除性质，值不参与高亮（CollectTerms 用）</summary>
            public bool ExcludeValue { get; set; }
            /// <summary>绑定后的显示列下标（rowData 下标，1 基）</summary>
            public int Index { get; set; } = -1;
            public override bool Evaluate(object[] rowData, ReadOnlySpan<bool> termHits)
            {
                if (Index < 1 || Index >= rowData.Length) return false;
                string text = CellValueAsText(rowData[Index]);
                return Op switch
                {
                    "" => text.Contains(Value, StringComparison.OrdinalIgnoreCase),
                    "==" => string.Equals(text.Trim(), Value.Trim(), StringComparison.OrdinalIgnoreCase),
                    "<>" => !text.Contains(Value, StringComparison.OrdinalIgnoreCase),
                    _ => TryParseNumber(text, out double cell)
                      && TryParseNumber(Value, out double target)
                      && (Op switch
                      {
                          ">" => cell > target,
                          "<" => cell < target,
                          ">=" => cell >= target,
                          "<=" => cell <= target,
                          _ => false,
                      }),
                };
            }
        }

        /// <summary>数值解析：接受千分位与前后空白；日期/带单位文本一律失败 → 该行不匹配</summary>
        internal static bool TryParseNumber(string text, out double value)
            => double.TryParse(text.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);

        /// <summary>解析产物：表达式树 + 去重整行词表（供逐行单趟匹配）</summary>
        public sealed class PreparedSearch
        {
            public SearchExpr Root { get; init; } = null!;
            public List<string> Terms { get; } = [];
            /// <summary>未绑定成功的列限定词（BindColumns 填充；供状态栏提示「无此列」）</summary>
            public List<string> UnresolvedColumns { get; } = [];
            /// <summary>正向高亮词（Term/列值；NotExpr 内的排除词不高亮）</summary>
            public List<string> HighlightTerms { get; } = [];
        }

        // ---------------- 解析 ----------------

        public PreparedSearch Prepare(string keyword)
        {
            string normalized = NormalizePunctuation(keyword);
            int pos = 0;
            var root = string.IsNullOrWhiteSpace(normalized)
                ? new TermExpr { Term = "" }
                : ParseOr(ref pos, normalized);

            var prepared = new PreparedSearch { Root = root };
            CollectTerms(root, prepared, negated: false);
            return prepared;
        }

        private string NormalizePunctuation(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;
            var sb = new StringBuilder(input.Length);
            foreach (char c in input)
            {
                switch (c)
                {
                    case '，': case ',': sb.Append(','); break;
                    case '；': case ';': sb.Append(';'); break;
                    case '！': case '!': sb.Append('!'); break;
                    case '＆': sb.Append('&'); break;
                    case '＋': sb.Append('+'); break;
                    case '＝': sb.Append('='); break;
                    case '（': sb.Append('('); break;
                    case '）': sb.Append(')'); break;
                    case '　': sb.Append(' '); break;
                    default: sb.Append(c); break;
                }
            }
            return sb.ToString();
        }

        private SearchExpr ParseOr(ref int pos, string input)
        {
            var left = ParseAnd(ref pos, input);
            while (pos < input.Length)
            {
                SkipSpaces(ref pos, input);
                if (pos >= input.Length) break;

                if (input[pos] == ',' || input[pos] == ';')
                {
                    pos++;
                    var right = ParseAnd(ref pos, input);
                    left = new OrExpr { Left = left, Right = right };
                }
                else break;
            }
            return left;
        }

        private SearchExpr ParseAnd(ref int pos, string input)
        {
            var left = ParseUnary(ref pos, input);
            while (pos < input.Length)
            {
                SkipSpaces(ref pos, input);
                if (pos >= input.Length) break;

                char c = input[pos];
                if (c == '&' || c == '+')
                {
                    pos++;
                    var right = ParseUnary(ref pos, input);
                    left = new AndExpr { Left = left, Right = right };
                }
                else break;
            }
            return left;
        }

        private SearchExpr ParseUnary(ref int pos, string input)
        {
            SkipSpaces(ref pos, input);
            if (pos >= input.Length) return new TermExpr { Term = "" };

            if (input[pos] == '!')
            {
                pos++;
                var child = ParseUnary(ref pos, input);
                return new NotExpr { Child = child };
            }
            return ParsePrimary(ref pos, input);
        }

        private SearchExpr ParsePrimary(ref int pos, string input)
        {
            SkipSpaces(ref pos, input);
            if (pos >= input.Length) return new TermExpr { Term = "" };

            if (input[pos] == '(')
            {
                pos++;
                var expr = ParseOr(ref pos, input);
                SkipSpaces(ref pos, input);
                if (pos < input.Length && input[pos] == ')') pos++;
                return expr;
            }
            return ParseTerm(ref pos, input);
        }

        private void SkipSpaces(ref int pos, string input)
        {
            while (pos < input.Length && input[pos] == ' ') pos++;
        }

        /// <summary>词终止字符（= > < 会把词切成「列限定」的左右两半）</summary>
        private static bool IsTermEnd(char c)
            => c is '&' or '+' or ',' or ';' or '(' or ')' or '=' or '>' or '<';

        /// <summary>列限定运算符首字符</summary>
        private static bool IsColumnOp(char c) => c is '=' or '>' or '<';

        private SearchExpr ParseTerm(ref int pos, string input)
        {
            int start = pos;
            while (pos < input.Length && !IsTermEnd(input[pos]))
                pos++;

            // 列限定语法：左词 <op> 右词，op ∈ = == > < >= <= <> !=
            // （左词非空才触发；裸运算符不作列限定，按普通词处理不致歧义）
            if (pos < input.Length && IsColumnOp(input[pos]) && pos > start)
            {
                string column = input[start..pos].Trim();
                char op0 = input[pos];
                pos++;

                // 左词以 ! 收尾 → != 不包含（8!=499 ≡ 8<>499 ≡ !8=499）
                bool bangEquals = column.EndsWith("!");
                if (bangEquals) column = column[..^1].TrimEnd();

                string op;
                if (!bangEquals && op0 == '<' && pos < input.Length && input[pos] == '>')
                {
                    op = "<>";   // 不包含
                    pos++;
                }
                else
                {
                    bool withEqual = pos < input.Length && input[pos] == '=';
                    op = (op0, withEqual) switch
                    {
                        ('=', false) => "",       // = 包含
                        ('=', true) => "==",      // == 整格等于
                        ('>', false) => ">",
                        ('>', true) => ">=",
                        ('<', false) => "<",
                        _ => "<=",
                    };
                    if (withEqual) pos++;
                }
                if (bangEquals) op = "<>";

                int valueStart = pos;
                while (pos < input.Length && !IsTermEnd(input[pos]))
                    pos++;
                string value = input[valueStart..pos].Trim();
                return new ColumnTermExpr { Column = column, Value = value, Op = op, ExcludeValue = op == "<>" };
            }

            string term = input[start..pos].Trim();
            return new TermExpr { Term = term };
        }

        // ---------------- 词表收集与列绑定 ----------------

        private static void CollectTerms(SearchExpr? expr, PreparedSearch target, bool negated)
        {
            switch (expr)
            {
                case TermExpr term:
                    if (string.IsNullOrWhiteSpace(term.Term)) return;   // 空词保持 Index=-1：匹配一切
                    int idx = target.Terms.FindIndex(x => string.Equals(x, term.Term, StringComparison.OrdinalIgnoreCase));
                    if (idx < 0) { target.Terms.Add(term.Term); idx = target.Terms.Count - 1; }
                    term.Index = idx;
                    if (!negated) AddHighlight(target, term.Term);
                    break;
                case AndExpr a: CollectTerms(a.Left, target, negated); CollectTerms(a.Right, target, negated); break;
                case OrExpr o: CollectTerms(o.Left, target, negated); CollectTerms(o.Right, target, negated); break;
                case NotExpr n: CollectTerms(n.Child, target, !negated); break;
                case ColumnTermExpr c:
                    // 不进整行词表（按列单独求值）；值作为正向词参与高亮，
                    // 排除性质的 <> / != 与被 ! 否定的列词不高亮
                    if (!negated && !c.ExcludeValue) AddHighlight(target, c.Value);
                    break;
            }
        }

        private static void AddHighlight(PreparedSearch target, string term)
        {
            if (term.Length == 0) return;
            if (!target.HighlightTerms.Contains(term, StringComparer.OrdinalIgnoreCase))
                target.HighlightTerms.Add(term);
        }

        /// <summary>
        /// 绑定列限定词的显示列下标。<paramref name="displayHeaders"/>[h] 对应显示列 h+1
        /// （即 rowData[h+1]），必须与界面表头同源（含 &lt;列N&gt; 与重名 _2 处理）；
        /// 传 null 时列限定仅支持列序号写法。
        /// </summary>
        public void BindColumns(PreparedSearch prepared, IReadOnlyList<string>? displayHeaders)
        {
            prepared.UnresolvedColumns.Clear();
            BindColumns(prepared.Root, displayHeaders, prepared);
        }

        private static void BindColumns(SearchExpr? expr, IReadOnlyList<string>? headers, PreparedSearch target)
        {
            switch (expr)
            {
                case ColumnTermExpr term: BindColumnTerm(term, headers, target); break;
                case AndExpr a: BindColumns(a.Left, headers, target); BindColumns(a.Right, headers, target); break;
                case OrExpr o: BindColumns(o.Left, headers, target); BindColumns(o.Right, headers, target); break;
                case NotExpr n: BindColumns(n.Child, headers, target); break;
            }
        }

        private static void BindColumnTerm(ColumnTermExpr term, IReadOnlyList<string>? headers, PreparedSearch target)
        {
            // 纯数字 = 1 基显示列序号；否则精确匹配表头名（OrdinalIgnoreCase）
            if (int.TryParse(term.Column, out int ordinal))
            {
                if (headers != null && (ordinal < 1 || ordinal > headers.Count))
                    target.UnresolvedColumns.Add($"列号{term.Column}");
                term.Index = ordinal;   // 越界由 Evaluate 的范围检查兜底
                return;
            }
            if (headers != null)
            {
                for (int h = 0; h < headers.Count; h++)
                {
                    if (string.Equals(headers[h], term.Column, StringComparison.OrdinalIgnoreCase))
                    {
                        term.Index = h + 1;
                        return;
                    }
                }
            }
            term.Index = -1;   // 无此列：永不匹配，并登记供状态栏提示
            target.UnresolvedColumns.Add(term.Column);
        }

        // ---------------- 逐行匹配 ----------------

        /// <summary>行匹配入口。<paramref name="prepared"/> 为 null = 不过滤（等价恒真）</summary>
        public static bool Matches(PreparedSearch? prepared, object[] rowData)
        {
            if (prepared == null) return true;

            var terms = prepared.Terms;
            Span<bool> hits = terms.Count <= 256 ? stackalloc bool[terms.Count] : new bool[terms.Count];
            MatchRowTerms(rowData, terms, hits);
            return prepared.Root.Evaluate(rowData, hits);
        }

        /// <summary>一趟扫行填充词命中表；全部命中即提前收工</summary>
        private static void MatchRowTerms(object[] rowData, List<string> terms, Span<bool> hits)
        {
            int remaining = terms.Count;
            for (int t = 0; t < terms.Count; t++) hits[t] = false;

            for (int i = 1; i < rowData.Length && remaining > 0; i++)
            {
                string text = CellValueAsText(rowData[i]);
                for (int t = 0; t < terms.Count; t++)
                {
                    if (!hits[t] && text.Contains(terms[t], StringComparison.OrdinalIgnoreCase))
                    {
                        hits[t] = true;
                        remaining--;
                    }
                }
            }
        }

        /// <summary>单元格值 → 匹配文本（唯一实现，解析与匹配共用；图片格取其文字部分，其余委托 <see cref="XlsxCellValue"/>，与显示同口径）</summary>
        internal static string CellValueAsText(object? value) => value switch
        {
            ImageTextPair p => p.Text ?? string.Empty,
            _ => XlsxCellValue.ToDisplayString(value)
        };
    }
}
