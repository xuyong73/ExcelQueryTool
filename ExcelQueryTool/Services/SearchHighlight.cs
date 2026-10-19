using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace ExcelQueryTool.Services
{
    /// <summary>
    /// 命中词高亮：把单元格文本按搜索词切分成 Run（命中片段加底色加粗）。
    ///
    /// 用法（见 UIStateManager 建列）：单元格 TextBlock 挂两个附加属性——
    /// <see cref="WatchTextProperty"/> 绑定列值、<see cref="WatchTermsProperty"/> 绑定
    /// <see cref="Instance.Terms"/>；任一变化触发重建。注意 TextBlock.Text 与 Inlines
    /// 是同一份内容存储：只在确认要输出高亮 Run 时才 Clear 重建，其余情况（无词表、
    /// 空文本、本格无命中）一律不动容器——否则清空后的内容无法自行恢复。
    ///
    /// 性能：只发生在可见行容器（虚拟化），重建为 O(文本长 × 词数) 的 IndexOf，
    /// 与搜索过滤路径完全无关。
    /// </summary>
    public sealed class SearchHighlight : INotifyPropertyChanged
    {
        public static SearchHighlight Instance { get; } = new();

        private List<string>? _terms;

        /// <summary>当前高亮词集（正向词，去重）；null/空 = 不高亮</summary>
        public List<string>? Terms
        {
            get => _terms;
            set { _terms = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Terms))); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        private static readonly Brush HighlightBrush = MakeBrush(0xFF, 0xE0, 0x8A);   // 柔和黄底
        private static readonly Brush HighlightForeground = MakeBrush(0x1F, 0x1F, 0x1F);

        private static Brush MakeBrush(byte r, byte g, byte b)
        {
            var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
            brush.Freeze();
            return brush;
        }

        /// <summary>单元格文本（通常与 TextBlock.Text 同源绑定），变化即重建 Inlines</summary>
        public static readonly DependencyProperty WatchTextProperty = DependencyProperty.RegisterAttached(
            "WatchText", typeof(object), typeof(SearchHighlight), new PropertyMetadata(null, OnWatchChanged));

        public static void SetWatchText(TextBlock element, object value) => element.SetValue(WatchTextProperty, value);
        public static object GetWatchText(TextBlock element) => element.GetValue(WatchTextProperty);

        /// <summary>高亮词表（绑定 Instance.Terms），变化即重建 Inlines</summary>
        public static readonly DependencyProperty WatchTermsProperty = DependencyProperty.RegisterAttached(
            "WatchTerms", typeof(object), typeof(SearchHighlight), new PropertyMetadata(null, OnWatchChanged));

        public static void SetWatchTerms(TextBlock element, object value) => element.SetValue(WatchTermsProperty, value);
        public static object GetWatchTerms(TextBlock element) => element.GetValue(WatchTermsProperty);

        private static void OnWatchChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is TextBlock tb) Rebuild(tb);
        }

        private static void Rebuild(TextBlock tb)
        {
            object? raw = tb.GetValue(WatchTextProperty);
            string text = raw as string ?? raw?.ToString() ?? "";
            var terms = Instance.Terms;

            // Text 与 Inlines 是同一份内容存储：不能无脑 Clear——清空会把 Text 绑定
            // 已写入的内容一并删掉，而 OneWay 绑定不因目标侧变化重传，单元格从此空白
            // （2026-09-24「过滤时部分单元格空白」的根因）。原则：只有确定要输出
            // 高亮 Run 时才 Clear 重建；无高亮输出时仅在残留旧高亮时收掉还原。
            List<(int Start, int End)> ranges = [];
            if (terms != null && text.Length > 0)
            {
                // 收集全部命中区间（允许词间重叠）
                foreach (string term in terms)
                {
                    if (term.Length == 0) continue;
                    int idx = 0;
                    while ((idx = text.IndexOf(term, idx, StringComparison.OrdinalIgnoreCase)) >= 0)
                    {
                        ranges.Add((idx, idx + term.Length));
                        idx++;
                    }
                }
            }

            if (ranges.Count == 0)
            {
                // 本格不高亮。若容器还挂着上一轮的旧高亮（词表变化/清空、词集不同、
                // 换行复用），收掉并还原纯文本；容器本就是纯文本/空则完全不动。
                if (HasHighlightRun(tb))
                {
                    tb.Inlines.Clear();
                    if (text.Length > 0) tb.Inlines.Add(new Run(text));
                }
                return;
            }

            ranges.Sort((a, b) => a.Start != b.Start ? a.Start.CompareTo(b.Start) : a.End.CompareTo(b.End));
            tb.Inlines.Clear();   // 走到这里必然马上重建高亮，清空是安全的

            int pos = 0;
            foreach (var (start, end) in ranges)
            {
                if (start < pos)
                {
                    if (end > pos) pos = end;   // 与已输出区间重叠：跳过，仅扩张右界
                    continue;
                }
                if (start > pos) tb.Inlines.Add(new Run(text[pos..start]));
                tb.Inlines.Add(new Run(text[start..end])
                {
                    Background = HighlightBrush,
                    Foreground = HighlightForeground,
                    FontWeight = FontWeights.Bold,
                });
                pos = end;
            }
            if (pos < text.Length) tb.Inlines.Add(new Run(text[pos..]));
        }

        /// <summary>容器里是否还挂着带底色的高亮 Run（旧词残留检测）</summary>
        private static bool HasHighlightRun(TextBlock tb)
        {
            foreach (var inline in tb.Inlines)
                if (inline is Run r && r.Background != null) return true;
            return false;
        }
    }
}
