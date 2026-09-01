using System.Text;

namespace ExcelQueryTool.Services
{
    public class SearchService
    {
        public abstract class SearchExpr
        {
            public abstract bool Evaluate(object[] rowData, Func<object[], string, bool> containsTerm);
        }

        public class AndExpr : SearchExpr
        {
            public SearchExpr Left { get; set; } = null!;
            public SearchExpr Right { get; set; } = null!;
            public override bool Evaluate(object[] rowData, Func<object[], string, bool> containsTerm) 
                => Left.Evaluate(rowData, containsTerm) && Right.Evaluate(rowData, containsTerm);
        }

        public class OrExpr : SearchExpr
        {
            public SearchExpr Left { get; set; } = null!;
            public SearchExpr Right { get; set; } = null!;
            public override bool Evaluate(object[] rowData, Func<object[], string, bool> containsTerm) 
                => Left.Evaluate(rowData, containsTerm) || Right.Evaluate(rowData, containsTerm);
        }

        public class NotExpr : SearchExpr
        {
            public SearchExpr Child { get; set; } = null!;
            public override bool Evaluate(object[] rowData, Func<object[], string, bool> containsTerm) 
                => !Child.Evaluate(rowData, containsTerm);
        }

        public class TermExpr : SearchExpr
        {
            public string Term { get; set; } = "";
            public bool Exclude { get; set; }
            public override bool Evaluate(object[] rowData, Func<object[], string, bool> containsTerm) 
                => Exclude ? !containsTerm(rowData, Term) : containsTerm(rowData, Term);
        }

        public SearchExpr ParseSearchConditions(string keyword)
        {
            string normalized = NormalizePunctuation(keyword);
            if (string.IsNullOrWhiteSpace(normalized)) return new TermExpr { Term = "", Exclude = false };

            int pos = 0;
            return ParseOr(ref pos, normalized);
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
            if (pos >= input.Length) return new TermExpr { Term = "", Exclude = false };

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
            if (pos >= input.Length) return new TermExpr { Term = "", Exclude = false };

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

        private SearchExpr ParseTerm(ref int pos, string input)
        {
            int start = pos;
            while (pos < input.Length && input[pos] != '&' && input[pos] != '+' && input[pos] != ',' && input[pos] != ';' && input[pos] != '(' && input[pos] != ')')
                pos++;

            string term = input.Substring(start, pos - start);
            if (string.IsNullOrEmpty(term)) return new TermExpr { Term = "", Exclude = false };

            term = term.Trim();
            if (term.StartsWith("!"))
            {
                string innerTerm = term.Substring(1).Trim();
                return new TermExpr { Term = innerTerm, Exclude = true };
            }
           return new TermExpr { Term = term, Exclude = false };
        }
    }
}