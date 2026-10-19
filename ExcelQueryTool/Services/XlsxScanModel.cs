using System.Buffers;

namespace ExcelQueryTool.Services;

/// <summary>工作表的结构化描述。</summary>
public sealed class SheetLayout
{
    /// <summary>列数</summary>
    public int ColumnCount { get; init; }

    /// <summary>
    /// 真实表头所在的物理行号（1 基）。数据行从 <c>HeaderRow + 1</c> 开始。
    /// 
    /// 不等于 1：很多表第 1 行是整表大标题（或整行缺失），列名在第 2/3 行。
    /// 调用方必须按本值跳过表头及其之前的行，不能硬编码「第 2 行起是数据」。
    /// </summary>
    public int HeaderRow { get; init; }

    /// <summary>各列表头原文（1 基，[0] 未使用）</summary>
    public string[] Headers { get; init; } = [];

    /// <summary>
    /// 各列在 xlsx 中声明的列宽（1 基列号作下标，[0] 未使用；字符单位，与 Excel 「列宽」一致）。
    /// 0 = 该列未显式声明（用默认宽度）。未解析到 &lt;col&gt; 时为空数组。
    /// </summary>
    public double[] ColumnWidths { get; init; } = [];

    /// <summary>
    /// 各列是否为「数字列」（1 基，[0] 未使用）。判定依据：该列数据区出现的
    /// 非空值中，数字占比 ≥ <see cref="NumericColumnClassifier.RatioPercent"/>/100
    /// （口径统一在 <see cref="NumericColumnClassifier"/>，与文本读取器共用）。
    ///
    /// 用途：列级对齐 —— Excel 的「常规」格式下数字右对齐、文本左对齐，
    /// 程序统一按此惯例渲染（不读 Excel 的 alignment，避免同列忽左忽右）。
    /// </summary>
    public bool[] ColumnIsNumeric { get; init; } = [];

    /// <summary>
    /// 本工作表「出现过非空单元格的最大列号」（1 基；0 = 未知/空表）。
    /// 由有界预扫（<c>ProbeMaxColumn</c>）得出，用途是把列数下限钉在真实数据宽度上 ——
    /// 表头行常比数据窄（尾部空表头未填样式），只按表头宽度定列会丢掉右侧数据。
    /// 
    /// 随布局一并传给数据加载路径，使其不必为同一个数字再扫一遍（见 <c>Scan</c> 的 knownLayout）。
    /// </summary>
    public int FullSheetMaxColumn { get; init; }
}

/// <summary>一行单元格的稀疏集合。可变对象，跨行复用以避免逐行分配。</summary>
public sealed class XlsxRowBuffer
{
    private int[] _cols = new int[64];
    private object?[] _vals = new object?[64];
    private int _count;

    public int Count => _count;
    public int ColumnAt(int i) => _cols[i];
    public object? ValueAt(int i) => _vals[i];

    /// <summary>
    /// <b>最后一个有值的列号</b>。只在 <see cref="Add"/> 时更新，因此空单元格
    /// （<c>&lt;c r="I1"/&gt;</c>，无 &lt;v&gt;）不计入。
    /// 不要把本属性当作「这一行的列数」用 —— 那是 <see cref="MaxColumnSeen"/>。
    /// </summary>
    public int MaxColumn { get; private set; }

    /// <summary>
    /// 本行<b>出现过的最大列号</b>（不论该单元格是否有值）。
    /// <c>&lt;c r="I1"/&gt;</c> 这种空单元格也计入，因此它才等于「这一行铺到了第几列」。
    /// 
    /// 与 <see cref="MaxColumn"/> 的区别是本项目的关键坑：很多表的表头行只有第一个格
    /// 有值（其余是带样式的空单元格），此时 MaxColumn=1 而 MaxColumnSeen=9。
    /// 判定列数必须用后者，否则会把整表压成 1 列、丢掉第 2 列起的全部数据。
    /// </summary>
    public int MaxColumnSeen { get; private set; }

    /// <summary>记录一个「出现过」的列号（空单元格也调用），供 <see cref="MaxColumnSeen"/> 累计</summary>
    internal void NoteColumnSeen(int column)
    {
        if (column > MaxColumnSeen) MaxColumnSeen = column;
    }

        internal void Reset()
        {
            // 只清引用，避免整片写入
            for (int i = 0; i < _count; i++)
                _vals[i] = null;
            _count = 0;
            MaxColumn = 0;
            MaxColumnSeen = 0;
        }

        internal void Add(int column, object? value)
        {
            if (_count >= _vals.Length)
            {
                int cap = _vals.Length * 2;
                Array.Resize(ref _cols, cap);
                Array.Resize(ref _vals, cap);
            }
            _cols[_count] = column;
            _vals[_count] = value;
            _count++;
            if (column > MaxColumn) MaxColumn = column;
        }

    /// <summary>按列号（1 基）取值，缺失返回 null</summary>
    public object? GetByColumn(int column)
    {
        for (int i = 0; i < _count; i++)
            if (_cols[i] == column) return _vals[i];
        return null;
    }
}

/// <summary>
/// 头部阶段截停后移交调用方的「尾部并行扫描上下文」。
///
/// <see cref="Buffer"/> 是从 <see cref="ArrayPool{Byte}.Shared"/> 租借的整块 sheet 部件字节
/// （头部阶段已扫过前缀，<see cref="TailStart"/> 起是未扫描的行对齐区间）；
/// <see cref="SharedStrings"/> / <see cref="DateStyles"/> 是只读共享部件，可被多个
/// worker 并发读取（扫描期间无人写入）。
/// 调用方用完必须 <see cref="Release"/> 归还缓冲。
/// </summary>
public sealed class ParallelTailContext
{
    public required byte[] Buffer { get; init; }
    public required int TailStart { get; init; }
    public required int Total { get; init; }
    /// <summary>尾区之前已出现过的 row 元素总数（顺序号种子，保证缺 r 行的行号与全量扫描一致）</summary>
    public required int RowsBeforeTail { get; init; }
    public required List<string> SharedStrings { get; init; }
    public required HashSet<int> DateStyles { get; init; }

    private readonly ArrayPool<byte> _pool = ArrayPool<byte>.Shared;

    public void Release() => _pool.Return(Buffer);
}
