using System.Buffers;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace ExcelQueryTool.Services;

/// <summary>
/// 自写的 xlsx 工作表流式读取器。
/// 
/// 为什么绕开 MiniExcel 的行映射：
/// MiniExcel.Query 每行产出 ExpandoObject，10 列时约 11 KB/行，
/// 100 万行实测约 10.8 GB 分配 / 10.8 秒。本读取器用 XmlReader 直接流式扫
/// sheetData，同文件实测约 605 MB / 3.7 秒（约 3 倍速度、18 倍分配优势）。
/// 
/// 行为约定：
/// - 表头行自动下探：第 1 行若是整表大标题（有值格数 ≤ <c>TitleRowMaxValuedCells</c>），
///   继续探第 2 行取列名（<c>HeaderProbeRows</c>）；数据从表头行之后开始
/// - 日期按单元格 style 的 numFmtId 判定（内置日期 ID 集合 + 自定义 numFmt 的 formatCode）
/// - 空单元格产出 null，显示为空串
/// </summary>
public sealed class XlsxSheetReader
{
    /// <summary>Excel 内置的日期/时间 numFmtId。ECMA-376 的 14~22 与 45~47；
    /// 27~36 / 50~58 是 locale 相关的日期格式（中文版 Excel/WPS 的「yyyy年m月」等走这些 ID），
    /// 不收录时这类日期会显示成 45707 之类的序列号。</summary>
    private static readonly HashSet<int> BuiltInDateFormats =
    [
        14, 15, 16, 17, 18, 19, 20, 21, 22,
        27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
        45, 46, 47,
        50, 51, 52, 53, 54, 55, 56, 57, 58,
    ];

    /// <summary>
    /// 表头下探的最大行数（2 行）。
    /// 只覆盖「第 1 行整表大标题 + 第 2 行列名」这一种真实结构 —— 实测所有带标题的业务表
    /// 都是这个形状。设成更大值反而危险：第 3 行通常已是数据，会把数据值当列名
    /// （实测「不良占比」sheet 探到第 3 行取到了「深圳工厂 / 1」）。
    /// </summary>
    private const int HeaderProbeRows = 2;

    /// <summary>
    /// 「首行算整表标题」的有值格数上限。首行有值格超过此值即认定它本身就是表头，不下探。
    /// 
    /// 取 2 的依据（实测）：标题行至多散布 1~3 个格（如「不良占比」的 B1/L1）；
    /// 而正常表头行远超此值 —— Loma FAI501_* 第 1 行有 17 个有值格，
    /// 它才是真正的表头（合并单元格让 MaxColumn 看起来只有 17），
    /// 第 2 行反而是参数区数据。不加本条件会把列数从 17 误涨到 67、表头换成数据值。
    /// </summary>
    private const int TitleRowMaxValuedCells = 2;

    /// <summary>
    /// 「整表最大列号」预扫的行数上限。
    /// 
    /// 取 10_000 与列占用抽样（<c>OccupancyProbeRows</c>）同量级，理由：两者都在回答
    /// 「表的真实宽度形状如何」这类只依赖表头附近的问题。实测（3 个样本文件共 277 张表）
    /// 前 10 行就与全表结论完全一致，取 1 万行留出极大余量，同时把成本从
    /// O(全表)（百万行文件 2.5 s）压到常数级。
    /// </summary>
    private const int MaxColumnProbeRows = 10_000;

    /// <summary>统计一行里有值的格数</summary>
    private static int CountValued(XlsxRowBuffer row)
    {
        int n = 0;
        for (int i = 0; i < row.Count; i++)
            if (IsMeaningfulValue(row.ValueAt(i))) n++;
        return n;
    }

    private readonly string _filePath;

    /// <summary>
    /// 本 reader 使用的包会话。默认自建（用完即弃），但调用方可通过
    /// <see cref="UseSession"/> 注入一个跨调用复用的会话 ——
    /// 见 <see cref="ExcelDataManager"/> 的会话池说明：挂在 reader 上无法跨调用复用。
    /// </summary>
    private XlsxPackageSession? _session;

    public XlsxSheetReader(string filePath) => _filePath = filePath;

    /// <summary>注入外部会话（所有权仍属调用方，本类不负责释放）</summary>
    internal XlsxSheetReader UseSession(XlsxPackageSession session)
    {
        _session = session;
        return this;
    }

    private XlsxPackageSession Session => _session ??= new XlsxPackageSession(_filePath);

    /// <summary>
    /// 最近一次 <see cref="Scan"/>（<c>stopAfterDataRows</c> &gt; 0 且走快路径且发生截停）产生的
    /// 尾部并行上下文；其余情况为 null。每次 <see cref="RunScan"/> 开头重置 ——
    /// 调用方取走后必须调用 <see cref="ParallelTailContext.Release"/> 归还缓冲。
    /// </summary>
    public ParallelTailContext? LastTailContext { get; private set; }

    /// <summary>行访问回调：返回 false 提前终止</summary>
    public delegate bool RowVisitor(int physicalRow, XlsxRowBuffer row);

    /// <summary>
    /// 单趟扫描入口：表头 → 逐行回调数据行。
    /// 
    /// 关键约束（也是本方法存在的理由）：调用方必须在<b>第一行数据之前</b>拿到最终列映射。
    /// 因此布局在首行数据到达时立即发布（见 <c>occupancyOnly</c> 的取舍说明），
    /// 空表头列的取舍采用<b>保守保留</b>，绝不丢行、列映射绝不随行漂移。
    /// 
    /// 行数据的投递保证：<paramref name="onLayout"/> 一旦返回 true，其后的<b>每一行</b>
    /// 都会传给 <paramref name="onRow"/>，不再丢行。
    /// </summary>
    /// <param name="onLayout">列映射确定时的回调；返回 false 表示不需要读数据行</param>
    /// <param name="onRow">数据行回调（第 2 行起）；返回 false 提前结束</param>
    /// <param name="knownLayout">
    /// 已在别处（如 <see cref="ReadLayoutWithOccupancy"/>）取得的布局。
    /// 
    /// 传入非 null 且列数大于 0 时，本方法<b>跳过整包重新打开</b>与 sharedStrings / styles
    /// 的重复解析，直接用它发布布局 —— UI 流程里列元数据与数据加载本就要各读一次包，
    /// 而每次开包都要重建 ZipArchive 中央目录并解析这些部件（实测 Loma 全书 268 表累计 2.5 s）。
    /// 
    /// 安全性由「两条路径共用同一套保守列映射」保证：
    /// 空表头列一律保留、显示列序 == 源列号（见 <see cref="ExcelDataManager"/> 的加载注释），
    /// 因此复用的布局与重新解析得到的布局必然一致。
    /// </param>
    public void Scan(
        string worksheetName,
        CancellationToken token,
        int occupancyProbeRows,
        Func<SheetLayout, bool> onLayout,
        RowVisitor onRow,
        SheetLayout? knownLayout = null,
        Func<string, HashSet<int>>? pictureColumnsProbe = null,
        int stopAfterDataRows = 0)
    {
        // 复用已知布局：省掉一次开包 + workbook/rels/sheet 定位 + 两个部件解析。
        // 注意此处仍要真正扫一遍数据行（行的值只在扫的时候才知道），
        // 省下的是「结构判定」那部分重复工作，不是行读取本身。
        if (knownLayout is { ColumnCount: > 0 } cached)
        {
            bool reuseStopped = false;
            if (!onLayout(cached)) return;

            int cachedHeaderRow = cached.HeaderRow;

            ReadLayoutAndRows(
                worksheetName,
                token,
                _ => { },
                (rowNo, row) =>
                {
                    if (reuseStopped) return false;
                    // 表头行（可能在第 2/3 行）不作为数据行投递
                    if (rowNo <= cachedHeaderRow) return true;
                    if (!onRow(rowNo, row)) { reuseStopped = true; return false; }
                    return true;
                },
                new ScanOptions
                {
                    EmitLayoutImmediately = true,
                    PictureColumnsProbe = pictureColumnsProbe,
                    // 布局里已记录整表最大列号：透传下去，省掉这次全表预扫
                    KnownFullSheetMaxColumn = cached.FullSheetMaxColumn,
                    StopAfterDataRows = stopAfterDataRows,
                });

            return;
        }

        bool delivered = false;
        bool stopRequested = false;
        int headerRowNo = 0;
        var layout = EmptyLayout;

        // 单趟：occupancyOnly 模式下，ReadLayoutAndRows 会在首行数据处发布布局，
        // 随后才开始投递数据行 —— 因此首行延迟不随占用抽样行数增长。
        ReadLayoutAndRows(
            worksheetName,
            token,
            _ => { },                                   // 表头回调：占用未统计完，忽略
            (rowNo, row) =>
            {
                if (stopRequested) return false;
                if (!delivered) { headerRowNo = rowNo; return true; }   // 布局未就绪 → 记表头行，不投递
                if (rowNo <= headerRowNo) return true;                 // 表头行（可能下探到第 2/3 行）
                if (!onRow(rowNo, row)) { stopRequested = true; return false; }
                return true;
            },
            new ScanOptions
            {
                RequireColumnOccupancy = true,
                OnLayoutComplete = l =>
                {
                    layout = l;
                    if (!delivered)
                    {
                        delivered = true;
                        stopRequested = !onLayout(l);
                    }
                },
                OccupancyRowLimit = occupancyProbeRows,
                OccupancyOnly = true,
                PictureColumnsProbe = pictureColumnsProbe,
                StopAfterDataRows = stopAfterDataRows,
            });

        // 工作表不存在 / 空表 / 布局未发布：兜底回调一次
        if (!delivered) onLayout(layout);
    }

    /// <summary>
    /// 工作簿内所有工作表名（保持原始顺序）。
    /// 复用调用方传入的包会话（不再各自开 FileStream + ZipArchive）；
    /// 包打不开时会话退化为「空包」，返回空列表。
    /// </summary>
    internal static List<string> GetSheetNames(XlsxPackageSession session)
        => session.RunLocked(zip =>
        {
            var names = new List<string>();
            foreach (var (name, _) in XlsxRelationships.EnumerateSheets(zip))
                names.Add(name);
            return names;
        }, new List<string>());

    /// <summary>
    /// 单趟遍历同时完成结构判定与数据产出（仅内部使用）。
    /// onLayout 在第 1 行解析后立即回调（默认），随后逐行调用 visitor。
    /// 扫描选项统一由 <see cref="ScanOptions"/> 承载（各字段默认值与原参数默认一致）。
    /// </summary>
    internal void ReadLayoutAndRows(
        string worksheetName,
        CancellationToken token,
        Action<SheetLayout> onLayout,
        RowVisitor visitor,
        ScanOptions options)
    {
        try
        {
            // 包与共享部件来自会话：同一文件内只开一次包、只解析一次
            // sharedStrings/styles（此前每次调用都重做，实测占单表成本 61%）。
            // 整段扫描在会话锁内执行 —— 扫描会持续打开部件流，不能与图片索引线程并发读同一个包。
            Session.RunLocked<object?>(zip =>
            {
                RunScan(zip, worksheetName, token, onLayout, visitor, options);
                return null;
            }, null);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch
        {
            // 解析失败时退化为「无数据」，不阻塞加载（图片/结构异常不应导致程序崩溃）
            onLayout(EmptyLayout);
        }
    }

    /// <summary>
    /// <see cref="ReadLayoutAndRows"/> 的扫描选项：历史上 12 个位置参数的收敛。
    /// 全部字段默认值与原参数默认一致（全关 / 0）。
    /// </summary>
    internal sealed class ScanOptions
    {
        /// <summary>
        /// 为 true 时即使 visitor 提前返回也会继续扫描数据区，直到占用抽样窗口
        /// （<see cref="OccupancyRowLimit"/> 行）用尽，保证「列级数字采样 / 结构判定」有足够样本；
        /// 为 false 时 visitor 一旦停止扫描立即结束。
        /// 该模式下 onLayout 仍在首行数据处回调（保证 visitor 可据此构建行数据）。
        /// </summary>
        public bool RequireColumnOccupancy { get; init; }

        /// <summary>与 onLayout 收到同一布局实例；用于结果回传</summary>
        public Action<SheetLayout>? OnLayoutComplete { get; init; }

        /// <summary>
        /// 列占用统计的最大数据行数（0 = 不限）。空列判定只需抽样即可：
        /// 前若干行已出现非空值的列即视为「有数据」，可避免为判定列集合而扫完整个大表。
        /// </summary>
        public int OccupancyRowLimit { get; init; }

        /// <summary>
        /// 为 true 时，在第一行数据到达前就发布布局（保证不丢行）。
        ///
        /// 为什么不在首行前等占用收敛：判定需要观察多行，而行缓冲跨行复用、无法补投 ——
        /// 「等观测足够再发布」必然丢掉发布前的所有行（实测 FAI* 表一行都投不出来）。
        /// 空表头列因此统一保守保留（显示为 &lt;列N&gt;），不影响数据正确性；
        /// 列级数字结论等采样窗口结束由 <see cref="ReflectNumericSamples"/> 回填。
        /// </summary>
        public bool OccupancyOnly { get; init; }

        /// <summary>
        /// 为 true 时，表头行解析后立即触发一次 onLayout，随后才投递数据行。
        /// 供「调用方已在别处拿到布局、只想扫数据行」的复用场景使用
        /// （见 <see cref="Scan"/> 的 <c>knownLayout</c> 参数）—— 此时不需要占用统计，
        /// 也不需要任何延迟，表头到手即发布可保证一行不丢。
        /// </summary>
        public bool EmitLayoutImmediately { get; init; }

        /// <summary>
        /// 可选：返回本工作表有图片的源列号（1 基）。用于把列数下限钉在图片锚点列上 ——
        /// 图片锚点常常落在表格数据区之外（如数据到第 8 列、图片锚在第 10 列），
        /// 若列数不含这些列，勾选「显示图片」也永远渲染不出来。
        /// 由调用方注入以避免本类依赖 <c>ImageManager</c>。
        /// </summary>
        public Func<string, HashSet<int>>? PictureColumnsProbe { get; init; }

        /// <summary>布局里已记录的整表最大列号：透传下去，省掉这次全表预扫</summary>
        public int KnownFullSheetMaxColumn { get; init; }

        /// <summary>数据行达到该数后停止（0 = 不限）</summary>
        public int StopAfterDataRows { get; init; }
    }

    /// <summary>扫描主体：在包会话锁内执行，<paramref name="zip"/> 由会话提供。</summary>
    private void RunScan(
        ZipArchive zip,
        string worksheetName,
        CancellationToken token,
        Action<SheetLayout> onLayout,
        RowVisitor visitor,
        ScanOptions options)
    {
        // 选项解构为局部变量：下文按原参数名引用，主体不动
        bool requireColumnOccupancy = options.RequireColumnOccupancy;
        var onLayoutComplete = options.OnLayoutComplete;
        int occupancyRowLimit = options.OccupancyRowLimit;
        bool occupancyOnly = options.OccupancyOnly;
        bool emitLayoutImmediately = options.EmitLayoutImmediately;
        var pictureColumnsProbe = options.PictureColumnsProbe;
        int knownFullSheetMaxColumn = options.KnownFullSheetMaxColumn;
        int stopAfterDataRows = options.StopAfterDataRows;

        LastTailContext = null;   // 每次扫描重置：旧尾部上下文要么已被调用方取走并释放，要么不存在

        {
            var (sharedStrings, dateStyles) = Session.SharedParts;

            string? sheetPart = Session.LocateSheet(worksheetName);
            if (sheetPart == null) { onLayout(EmptyLayout); return; }

            bool layoutSent = false;          // 布局是否已回调（含 onLayoutComplete）
            int totalCols = 0;
            string[]? headers = null;
            bool visitorStopped = false;
            int dataRowSeen = 0;
            int headerRowNo = 0;              // 真实表头所在物理行号（0 = 尚未确定）
            int scannedRowNo = 0;             // 已扫过的行计数（含空行）
            int firstRowNo = 0;               // 首个出现的物理行号
            int firstRowValued = 0;           // 首行最后有值列
            int firstRowValuedCells = 0;      // 首行有值格数
            int firstRowSeen = 0;             // 首行铺开列数

            // 数字列判定采样（列级对齐用）：只采前若干行，避免为对齐多扫一遍全表。
            // 计数「非空值个数」与「其中数字个数」，扫完按占比定案
            // （占比/样本口径统一在 NumericColumnClassifier，与文本读取器共用）。
            int[]? numericSample = null;    // 每列非空值数
            int[]? numericHit = null;       // 每列数字值数
            int numericSampleRows = 0;
            bool[]? numericConclusion = null;   // 最终结论（扫完后回填）

            // 全表最大列号：有界预扫求得（只求列宽，纯 XmlReader 计数，不取值、不建对象）。
            // 用途：把 totalCols 的下限钉在真实数据宽度与图片锚点列上 ——
            // 表头行的 MaxColumnSeen 只反映「表头那几行铺到第几列」，
            // 若表头行较窄（或整行缺失）而数据更宽，仅靠它会压窄列数、丢掉后段列。
            // 
            // 复用调用方已有布局时直接取它记录的值，不再重扫（见 knownLayout 参数）。
            // 取 > 0 判据而非 ??：0 是「未知/空表」的合法取值，此时做一次有界预扫代价很低，
            // 换来的是「宁可按更宽定列、也不丢右侧数据」——与列数取 max 的整体取向一致。
            int fullSheetMaxCol = knownFullSheetMaxColumn > 0
                ? knownFullSheetMaxColumn
                : ProbeMaxColumn(zip, sheetPart, token, MaxColumnProbeRows);
            var columnWidths = LoadColumnWidths(zip, sheetPart, token);
            int pictureMaxCol = 0;
            foreach (int pc in pictureColumnsProbe?.Invoke(worksheetName) ?? [])
                if (pc > pictureMaxCol) pictureMaxCol = pc;

            // occupancyOnly 模式：不提前发布（要在首行数据处发布，见数据行分支注释）。
            // emitLayoutImmediately（复用布局场景）：表头到手即发布。
            bool publishAtHeader = emitLayoutImmediately && !occupancyOnly;

            // 当前布局快照（供两处回调共用：onLayout 与 onLayoutComplete 收到同一实例）
            SheetLayout Snapshot() => new()
            {
                ColumnCount = totalCols,
                HeaderRow = headerRowNo,
                Headers = headers ?? [],
                ColumnWidths = columnWidths,
                ColumnIsNumeric = numericConclusion ?? [],
                FullSheetMaxColumn = fullSheetMaxCol,
            };

            void PublishLayout()
            {
                if (layoutSent) return;
                layoutSent = true;
                var snapshot = Snapshot();
                onLayout(snapshot);
                onLayoutComplete?.Invoke(snapshot);
            }

            // 采样到的「非空值数 / 数字值数」→ 列级数字列结论（就地写入 numericConclusion，
            // 与 Snapshot 共用同一数组引用，故回填后已发布的布局也能拿到结果）
            void ReflectNumericSamples()
            {
                if (numericSample == null || numericHit == null || totalCols <= 0) return;

                numericConclusion = new bool[totalCols + 1];
                for (int c = 1; c <= totalCols; c++)
                    numericConclusion[c] = NumericColumnClassifier.Classify(numericHit[c], numericSample[c]);
            }

            // 头部截停计数（并行尾巴用）：只数「已投给 visitor 的数据行」，
            // 与 dataRowSeen（所有遇到的数据行，含 visitorStopped 后的）口径不同。
            int stopAfterSeen = 0;
            byte[]? tailBuffer = null;
            int tailStart = 0, tailTotal = 0;

            ScanRows(zip, sheetPart, sharedStrings, dateStyles, token, (rowNo, row) =>
            {
                scannedRowNo++;

                // ── 表头行判定 ──
                // 不再硬编码「物理第 1 行」：真实业务表常见「第 1 行是整表大标题
                // （或整行缺失），第 2 行才是列名」。原先只看第 1 行会让列数被压到 1、
                // 甚至整表判空（实测客户投诉 / OSS / P25 / 越南 BZ5 四表全废）。
                // 
                // 判定策略（保守，绝不比原行为更差）：
                //   1. 第 1 行有值格 > 2 → 它本身就是表头，直接用（既有行为，零变更）。
                //      实测 Loma FAI501_* 第 1 行有 17 个有值格（合并单元格让 MaxColumn 偏小），
                //      第 2 行却是参数区数据；不加此条件会把表头换成数据值、列数 17→67。
                //   2. 第 1 行有值格 ≤ 2（明显是标题）且第 2 行「像列名」→ 用第 2 行。
                //   3. 其余情况仍用第 1 行。
                if (headerRowNo == 0)
                {
                    int seen = row.MaxColumnSeen;      // 本行铺到第几列（含空格）
                    int valued = row.MaxColumn;        // 本行最后一个有值的列
                    int valuedCells = CountValued(row); // 本行有值的格数

                    bool firstSeenRow = scannedRowNo == 1;

                    if (firstSeenRow)
                    {
                        firstRowNo = rowNo;
                        firstRowValued = valued;
                        firstRowValuedCells = valuedCells;
                        firstRowSeen = seen;

                        // 第 1 行有值格多 → 它本身就是表头，立即定案
                        if (valuedCells > TitleRowMaxValuedCells)
                        {
                            headerRowNo = rowNo;
                        }
                        else if (scannedRowNo >= HeaderProbeRows)
                        {
                            // 窗口已满且无法下探（单行表）：用首行
                            headerRowNo = firstRowNo;
                        }
                    }
                    else
                    {
                        // 第 2 行：仅当第 1 行「像标题」且本行「像列名」时才取代它
                        bool titleLike = firstRowValuedCells <= TitleRowMaxValuedCells;
                        bool headerLike = valuedCells >= 2
                                          && valuedCells * 3 >= Math.Max(seen, 1)
                                          && (valued > firstRowValued || valuedCells > firstRowValuedCells);

                        if (titleLike && headerLike)
                        {
                            headerRowNo = rowNo;
                            seen = Math.Max(seen, firstRowSeen);
                        }
                        else
                        {
                            // 下探失败：回退首行。首行的值已随行缓冲复用被覆盖、取不回来，
                            // 因此这里只能沿用「首行即表头」的宽度结论（用首行 seen 兜底）。
                            headerRowNo = firstRowNo;
                            seen = Math.Max(seen, firstRowSeen);
                        }
                    }

                    if (headerRowNo != 0)
                    {
                        // 列数取三者最大：表头行铺开的列、全表实际最大列、图片锚点最大列。
                        // 取 max 而非只看表头：表头行常比数据窄（尾部空表头未填样式），
                        // 只按表头宽度定列会丢掉右侧数据；图片列同理（锚点常在数据区之外）。
                        totalCols = Math.Max(Math.Max(seen, fullSheetMaxCol), pictureMaxCol);
                        if (totalCols <= 0) return false;      // 空表

                        headers = new string[totalCols + 1];

                        // 数字列判定采样器：C 列计「非空值数 / 数字值数」
                        numericSample = new int[totalCols + 1];
                        numericHit = new int[totalCols + 1];

                        for (int i = 0; i < row.Count; i++)
                        {
                            int col = row.ColumnAt(i);
                            if (col < 1 || col > totalCols) continue;
                            headers[col] = XlsxCellValue.ToDisplayString(row.ValueAt(i));
                        }

                        // 表头到手即可发布（复用布局 / 普通模式）：结构已可用，
                        // 空表头列由调用方按保守规则处理（一律保留）。
                        if (publishAtHeader) PublishLayout();
                    }
                    // 表头尚未确定：本行是标题/空行，不作为数据行投递
                    return true;
                }

                // 表头行本身不作为数据行投递（与原先「第 2 行起才是数据」的语义一致）
                if (rowNo <= headerRowNo) return true;

                // ── 数据行 ──
                // 数字列采样（前 NumericColumnClassifier.SampleRows 行）：统计每列非空值中数字的占比
                if (numericSampleRows < NumericColumnClassifier.SampleRows && numericSample != null && numericHit != null)
                {
                    numericSampleRows++;
                    for (int i = 0; i < row.Count; i++)
                    {
                        int col = row.ColumnAt(i);
                        if (col < 1 || col > totalCols) continue;

                        object? v = row.ValueAt(i);
                        if (!IsMeaningfulValue(v)) continue;

                        numericSample[col]++;
                        if (IsNumericValue(v)) numericHit[col]++;
                    }
                }

                // 数据行计数：驱动「元数据扫描抽样窗口用尽即停」（occupancyRowLimit）。
                // 此前兼喂列占用统计（hasData → ColumnHasData），该结论无消费方后
                // 只剩本职责；requireColumnOccupancy=false 的路径不设 occupancyRowLimit，
                // 多计的数不会触发任何提前停，行为不变。
                dataRowSeen++;

                // occupancyOnly：必须在第一行数据进入 visitor 之前发布布局。
                // 原因：visitor（Scan 的包装）在布局就绪前会丢弃行，而行缓冲跨行复用、
                // 补不回来 —— 任何「等观测够了再发布」的延迟都等价于丢行
                // （实测：FAI* 表因空表头列迟迟未判定，一行都投不出来）。
                // 因此这里不做收敛判定，首行即发布；空表头列的取舍改由调用方按
                // 「保守保留」处理（见 occupancyOnly 参数说明）。
                if (!layoutSent && totalCols > 0)
                {
                    PublishLayout();
                }

                // 占用抽样窗口已满、且调用方已不再需要数据行 → 无需继续读下去。
                // 
                // 这是「元数据扫描为什么也慢」的根因：ReadLayoutWithOccupancy 的 visitor
                // 恒返回 false（它不要数据行），原实现在此处仅置 visitorStopped 后继续扫，
                // 于是百万行文件被完整读一遍（实测 ~4.1 s），而占用统计只采前 1 万行、
                // 后续行的读取结果全部丢弃。抽样窗口用尽即停，结论一字不变。
                if (visitorStopped && occupancyRowLimit > 0 && dataRowSeen >= occupancyRowLimit)
                    return false;

                if (visitorStopped) return requireColumnOccupancy;

                stopAfterSeen++;
                if (!visitor(rowNo, row))
                {
                    visitorStopped = true;
                    // 抽样已够且不再需要统计：可提前结束
                    if (occupancyRowLimit > 0 && dataRowSeen >= occupancyRowLimit) return false;
                    return requireColumnOccupancy;
                }

                // 头部截停（并行尾巴）：本行刚已投出，停在它的结束标签之后 ——
                // 尾部从「未投出的第一行」开始，绝不丢行、绝不双投
                if (stopAfterDataRows > 0 && stopAfterSeen >= stopAfterDataRows)
                    return false;
                return true;
            }, stopAfterDataRows, out tailBuffer, out tailStart, out tailTotal);

            // 快路径头部截停：把尾部字节区间移交调用方（并行解析）。
            // XmlReader 回退路径不产生尾部（整表已由本趟扫完，行为与从前完全一致）。
            if (tailBuffer != null)
            {
                LastTailContext = new ParallelTailContext
                {
                    Buffer = tailBuffer,
                    TailStart = tailStart,
                    Total = tailTotal,
                    RowsBeforeTail = scannedRowNo,
                    SharedStrings = sharedStrings,
                    DateStyles = dateStyles,
                };
            }

            // 采样结束：数字列结论落成最终数组。
            // 必须在收尾发布之前调用 —— 循环内发布的快照与这个数组共享引用，
            // 回填后即生效；但收尾这次 onLayoutComplete 是新建数组快照的时机，
            // 早于它回填才能拿到结论。
            ReflectNumericSamples();

            // 收尾：布局从未发布过（空表 / 工作表不存在 / 只有表头无数据行 /
            // 空表头列始终没观测到数据 —— 按「整列为空」定案）
            if (!layoutSent)
            {
                if (totalCols <= 0) { onLayout(EmptyLayout); return; }
                PublishLayout();
            }

            // 扫完后追加一次回调（非 occupancyOnly 路径；
            // occupancyOnly 下 PublishLayout 已一并触发过 onLayoutComplete）
            if (requireColumnOccupancy && headers != null && onLayoutComplete != null && !occupancyOnly)
            {
                onLayoutComplete(Snapshot());
            }
        }
    }

    /// <summary>值是否算作「有数据」（非 null、非空串）</summary>
    private static bool IsMeaningfulValue(object? value)
    {
        if (value == null) return false;
        if (value is string s) return s.Length > 0;
        return true;
    }

    /// <summary>
    /// 只做结构判定：读表头 + 列宽 + 有界数字列采样（供 UI 列级对齐）。
    ///
    /// 抽样上限 <paramref name="probeRows"/> 行，扫完即停，因此不会随数据量变慢。
    /// UI 建列流程用它一次取得 Headers / ColumnWidths / ColumnIsNumeric / FullSheetMaxColumn，
    /// 布局随后透传给 <see cref="Scan"/> 的 knownLayout 复用。
    /// </summary>
    public SheetLayout ReadLayoutWithOccupancy(
        string worksheetName,
        CancellationToken token,
        int probeRows,
        Func<string, HashSet<int>>? pictureColumnsProbe = null)
    {
        SheetLayout result = EmptyLayout;

        ReadLayoutAndRows(
            worksheetName,
            token,
            _ => { },                                   // 立即回调的结构先忽略，以 complete 版为准
            (_, _) => false,                            // 不取数据行
            new ScanOptions
            {
                RequireColumnOccupancy = true,
                OnLayoutComplete = l => result = l,
                OccupancyRowLimit = probeRows,
                PictureColumnsProbe = pictureColumnsProbe,
            });

        return result;
    }

    private static SheetLayout EmptyLayout => new() { ColumnCount = 0 };

    /// <summary>
    /// 预扫 sheetData，只求<b>整表出现过的最大列号</b>。
    /// 
    /// 为什么需要它：列数若只按表头行定，表头窄（或整行缺失）时会把表压窄、丢掉右侧数据；
    /// 图片锚点落在数据区外时也需要把列数放宽到锚点列。两者都要求知道「全表铺到第几列」。
    /// 
    /// 成本：纯 <see cref="XmlReader"/> 顺序扫，只读 &lt;c&gt; 的 r 属性取前导字母，
    /// 不解析共享串、不构造值对象、无堆分配。
    /// 
    /// <para>
    /// <b>有界扫描（<paramref name="rowLimit"/>）</b>：原本这里扫完全表，但实测
    /// 百万行文件要 2.5 s —— 而「最大列号」这个结论在前几行就已经稳定：
    /// 全部 3 个真实样本文件共 277 张表逐一比对，<b>前 10 行 / 前 500 行 / 全表三者结果完全一致</b>
    /// （行是连续的、最宽的行总在表头附近）。因此改为扫到第 <paramref name="rowLimit"/> 行即停，
    /// 结论不变而成本从 O(全表) 降到 O(1)。
    /// 
    /// 传 0 或不大于 0 表示不限（保留原全表行为）。
    /// </para>
    /// </summary>
    private static int ProbeMaxColumn(ZipArchive zip, string sheetPart, CancellationToken token, int rowLimit = 0)
    {
        var entry = zip.GetEntry(sheetPart);
        if (entry == null) return 0;

        int maxCol = 0;
        bool inRow = false;
        int rowMax = 0;
        int rowsSeen = 0;
        long readCounter = 0;
        bool stopped = false;

        try
        {
            if (zip.OpenPartReader(sheetPart, XlsxXml.BufferSheet) is not { } reader) return 0;
            using (reader)
            while (!stopped && reader.Read())
            {
                if ((++readCounter & 0x1FFF) == 0) token.ThrowIfCancellationRequested();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "row":
                            if (rowLimit > 0 && rowsSeen >= rowLimit) { stopped = true; break; }
                            rowsSeen++;
                            inRow = true;
                            rowMax = 0;
                            break;

                        case "c":
                            if (!inRow || !reader.HasAttributes) break;
                            for (int ai = 0; ai < reader.AttributeCount; ai++)
                            {
                                reader.MoveToAttribute(ai);
                                if (reader.LocalName == "r")
                                {
                                    int c = XlsxValueParsing.ParseColumnLetters(reader.Value);
                                    if (c > rowMax) rowMax = c;
                                }
                            }
                            reader.MoveToElement();
                            break;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row")
                {
                    if (rowMax > maxCol) maxCol = rowMax;
                    inRow = false;
                }
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch
        {
            // 预扫失败不致命：退化为「不参与列数下限」，由表头行宽度兜底
            return 0;
        }

        return maxCol;
    }

    /// <summary>
    /// 流式扫 sheetData。physicalRow 为 Excel 实际行号（1 基）。
    /// visitor 返回 false 时提前结束。
    /// </summary>
    private static void ScanRows(
        ZipArchive zip,
        string sheetPart,
        List<string> sharedStrings,
        HashSet<int> dateStyles,
        CancellationToken token,
        RowVisitor visitor,
        int stopAfterDataRows,
        out byte[]? tailBuffer,
        out int tailStart,
        out int tailTotal)
    {
        tailBuffer = null;
        tailStart = 0;
        tailTotal = 0;

        var entry = zip.GetEntry(sheetPart);
        if (entry == null) return;

        // ── 快路径：手写字节级扫描器 ──
        // 把工作表部件读成字节，在 UTF-8 字节上直接定位标签/属性，避免 XmlReader
        // 为 3775 万个节点逐次物化属性字符串（实测解析段 3436 → 625 ms）。
        // 语义由 XlsxByteScanner 逐项镜像本方法；任何不支持的结构返回 false，
        // 立刻回退到下面的 XmlReader 路径（开新的流，与原行为完全一致）。
        if (TryScanFast(entry, sharedStrings, dateStyles, token, visitor,
                        stopAfterDataRows, out tailBuffer, out tailStart, out tailTotal)) return;

        var row = new XlsxRowBuffer();
        int currentColumn = 1;
        int sequentialColumn = 1;
        string? cellType = null;
        bool cellIsDate = false;
        bool cellHasValue = false;
        bool inCell = false;
        bool inValue = false;
        bool inInlineString = false;
        int cellVm = 0;
        bool stop = false;
        int currentRowNumber = 0;
        int sequentialRow = 0;
        var inlineText = new StringBuilder(64);
        long readCounter = 0;

        if (zip.OpenPartReader(sheetPart, XlsxXml.BufferSheet) is not { } reader) return;   // sheet 部件可能数百 MB，用大缓冲
        using (reader)
        while (!stop && reader.Read())
        {
            if ((++readCounter & 0x7FF) == 0) token.ThrowIfCancellationRequested();

            switch (reader.NodeType)
            {
                case XmlNodeType.Element:
                    switch (reader.LocalName)
                    {
                        case "row":
                            row.Reset();
                            sequentialColumn = 1;
                            sequentialRow++;
                            currentRowNumber = ParseRowNumber(reader, sequentialRow);
                            break;

                        case "c":
                            inCell = true;
                            cellHasValue = false;
                            cellType = null;
                            cellIsDate = false;
                            cellVm = 0;
                            inlineText.Clear();

                            // 热路径：每个 <c> 都要读 t / s / r 三个属性。
                            // 不用 GetAttribute(name) 逐个按名查找（三次名字比较），改为一次
                            // 顺序遍历 —— OOXML 里属性少（通常 0~3 个），顺序扫描更快且零分配。
                            // 属性顺序在 XML 中无保证，因此按 LocalName 分派而非按位置。
                            if (reader.HasAttributes)
                            {
                                for (int ai = 0; ai < reader.AttributeCount; ai++)
                                {
                                    reader.MoveToAttribute(ai);
                                    switch (reader.LocalName)
                                    {
                                        case "t":
                                            cellType = reader.Value;
                                            break;
                                        case "vm":
                                            // vm="N"：Excel 新版内嵌图片（richValue）的元数据索引。
                                            // 这类单元格 t="e"、值恒为 #VALUE!，真实内容是图片。
                                            if (XlsxValueParsing.TryParseNonNegativeInt(reader.Value, out int vmIdx))
                                                cellVm = vmIdx + 1;   // 存 1 基，0 = 无
                                            break;
                                        case "s":
                                            // 手写小整数解析：样式索引恒为非负整数，避免通用解析器开销
                                            if (XlsxValueParsing.TryParseNonNegativeInt(reader.Value, out int styleIdx))
                                                cellIsDate = dateStyles.Contains(styleIdx);
                                            break;
                                        case "r":
                                            // r 形如 "C12"：只取前导字母部分直接算列号，不切片、不分配、
                                            // 不做 ToUpperInvariant（原实现每格为此分配一个临时字符串）。
                                            int col = XlsxValueParsing.ParseColumnLetters(reader.Value);
                                            currentColumn = col > 0 ? col : sequentialColumn;
                                            break;
                                    }
                                }
                                reader.MoveToElement();
                            }
                            else
                            {
                                currentColumn = sequentialColumn;
                            }

                            // 记录「本行铺到第几列」。必须在 Add 之外单独记：
                            // 空单元格（<c r="I1"/> 无 <v>）从不 Add，若只靠 MaxColumn，
                            // 表头行有值格仅第 1 列时会把列数误判为 1（真实踩坑，见 MaxColumnSeen 注释）。
                            row.NoteColumnSeen(currentColumn);
                            break;

                        case "v":
                            inValue = true;
                            break;

                        case "is":
                            inInlineString = true;
                            inlineText.Clear();
                            break;

                        case "t":
                            // 仅在 <is> 内的 <t> 才是内联文本；消费后读取器会停在 </c> 或下一个兄弟节点，
                            // 因此必须在本次消费内完成状态收尾（下方 EndElement 不会再收到 "is"）。
                            if (inInlineString && !reader.IsEmptyElement)
                            {
                                inlineText.Append(reader.ReadElementContentAsString());
                                inInlineString = false;
                            }
                            break;
                    }
                    break;

                case XmlNodeType.Text:
                case XmlNodeType.SignificantWhitespace:
                    if (inValue && !inInlineString)
                    {
                        // vm（richValue 元数据）单元格：t="e" 值恒为 #VALUE!，真实内容是
                        // 内嵌图片（由 XlsxPictureIndex 按 (行,列) 另行解析），这里不投错误文本
                        if (cellVm == 0)
                        {
                            // 与字节扫描路径共用同一份值解析（共享字符串查表 / 日期样式转 DateTime 在此完成）
                            row.Add(currentColumn, XlsxValueParsing.ResolveCellType(reader.Value, cellType, cellIsDate, sharedStrings));
                            sequentialColumn = currentColumn + 1;
                            cellHasValue = true;
                        }
                        // 注意：此处不重置 inCell —— 单元格状态的唯一收敛点是
                        // EndElement: c（下方），否则那段兜底重置对「有值单元格」永远不可达。
                    }
                    break;

                case XmlNodeType.EndElement:
                    switch (reader.LocalName)
                    {
                        case "v":
                            inValue = false;
                            break;

                        case "c":
                            // 单元格状态的唯一收敛点：无论本格是否有值、是内联串还是异常结构，
                            // 都在这里清干净，确保不残留到下一个 <c>。
                            if (inCell)
                            {
                                // 内联文本已由 <t> 分支消费并暂存，此处落格
                                if (!cellHasValue && inlineText.Length > 0)
                                {
                                    row.Add(currentColumn, inlineText.ToString());
                                    sequentialColumn = currentColumn + 1;
                                }
                            }
                            inCell = false;
                            cellHasValue = false;
                            cellType = null;
                            cellIsDate = false;
                            inlineText.Clear();
                            inInlineString = false;
                            inValue = false;
                            break;

                        case "row":
                            if (!visitor(currentRowNumber, row)) stop = true;
                            break;
                    }
                    break;
            }
        }
    }

    /// <summary>
    /// 快路径：把 sheet 部件整体读入字节缓冲，交给 <see cref="XlsxByteScanner"/> 扫描。
    ///
    /// 读整个部件而非流式读，是为了让扫描器能在连续字节上做前后查找（定位标签结束、
    /// 属性引号配对）。实测 sheet1.xml 解压后 319 MB —— 单次分配可接受，
    /// 换来的是解析段 5.5× 提速；用完立即归还缓冲池。
    ///
    /// <paramref name="stopAfterDataRows"/> &gt; 0 且 visitor 据此提前截停时，
    /// 字节缓冲的所有权经 <paramref name="tailBuffer"/> 移交调用方（并行解析尾部），
    /// 由调用方负责归还缓冲池；此时本方法不再归还。
    ///
    /// 返回 false 表示「不支持该输入」（无 sheetData / UTF-16 / 取消以外的异常），
    /// 调用方须回退到 XmlReader 路径。取消异常照常向上抛。
    /// </summary>
    private static bool TryScanFast(
        ZipArchiveEntry entry,
        List<string> sharedStrings,
        HashSet<int> dateStyles,
        CancellationToken token,
        RowVisitor visitor,
        int stopAfterDataRows,
        out byte[]? tailBuffer,
        out int tailStart,
        out int tailTotal)
    {
        tailBuffer = null;
        tailStart = 0;
        tailTotal = 0;

        try
        {
            long len = entry.Length;
            if (len <= 0) return false;

            using var stream = entry.Open();

            // 大于 2 GB 的部件不作为：单数组长度上限（sheet 实体不会到这一步）
            if (len > int.MaxValue) return false;

            var pool = ArrayPool<byte>.Shared;
            byte[] buffer = pool.Rent((int)len);
            bool owned = true;   // 缓冲归还责任：true=本方法归还，false=已移交调用方
            try
            {
                int total = 0;
                while (total < len)
                {
                    int read = stream.Read(buffer, total, (int)(len - total));
                    if (read <= 0) break;
                    total += read;
                }
                if (total <= 0) return false;

                // visitor 需要原始委托类型；扫描器的回调签名同形，做一次薄包装
                bool ok = XlsxByteScanner.TryScan(
                    buffer.AsSpan(0, total), sharedStrings, dateStyles, token,
                    (rowNo, row) => visitor(rowNo, row), out int endOffset);

                if (ok && stopAfterDataRows > 0 && endOffset < total)
                {
                    // 头部截停：缓冲移交（调用方用完 ParallelTailContext.Release() 归还）
                    tailBuffer = buffer;
                    tailStart = endOffset;
                    tailTotal = total;
                    owned = false;
                }

                return ok;
            }
            finally
            {
                if (owned) pool.Return(buffer);
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (System.Xml.XmlException)
        {
            // 畸形 XML（含未定义实体、非法字符引用）：与 XmlReader 路径同样把
            // XmlException 冒给上层，保证两条路径对坏文件的报错行为一致。
            // 这里绝不能静默回退——本扫描器可能已向 visitor 投出行，回退重扫会双投。
            throw;
        }
        catch
        {
            // 不支持的输入：交给 XmlReader 路径重试（原行为）。
            // 注意：仅对「尚未投出任何行」的早期失败安全（BOM、超大部件、读流失败等）。
            return false;
        }
    }

    /// <summary>
    /// 解析工作表 &lt;cols&gt; 里声明的列宽（字符单位，与 Excel「列宽」数值一致）。
    /// 解析失败时返回空数组，不致命。
    /// </summary>
    private static double[] LoadColumnWidths(ZipArchive zip, string sheetPart, CancellationToken token)
    {
        // <col min max width customWidth?>：min/max 是范围，先按稀疏 (start,end,width) 收集
        var ranges = new List<(int Start, int End, double Width)>();

        try
        {
            if (zip.OpenPartReader(sheetPart, XlsxXml.BufferMedium) is not { } reader) return [];
            using (reader)

            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.LocalName == "sheetData") break;   // cols 一定在 sheetData 之前
                if (reader.LocalName != "col") continue;

                double width = 0;
                int min = 0, max = 0;
                for (int ai = 0; ai < reader.AttributeCount; ai++)
                {
                    reader.MoveToAttribute(ai);
                    switch (reader.LocalName)
                    {
                        case "min": min = XlsxValueParsing.TryParseNonNegativeInt(reader.Value, out int mn) ? mn : 0; break;
                        case "max": max = XlsxValueParsing.TryParseNonNegativeInt(reader.Value, out int mx) ? mx : 0; break;
                        case "width": double.TryParse(reader.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out width); break;
                    }
                }
                reader.MoveToElement();

                if (min < 1 || width <= 0) continue;
                if (max < min) max = min;
                ranges.Add((min, max, width));
            }

            if (ranges.Count == 0) return [];

            int maxCol = ranges.Max(r => r.End);
            var widths = new double[maxCol + 1];
            foreach (var (start, end, w) in ranges)
                for (int c = start; c <= end && c < widths.Length; c++)
                    widths[c] = w;

            return widths;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch
        {
            return [];
        }
    }

    private static int ParseRowNumber(XmlReader reader, int fallback)
    {
        string? r = reader.GetAttribute("r");
        return r != null && XlsxValueParsing.TryParseNonNegativeInt(r, out int n) ? n : fallback;
    }

    /// <summary>
    /// 读取 xl/sharedStrings.xml → 共享字符串数组（下标即 &lt;c t="s"&gt;&lt;v&gt;N&lt;/v&gt; 的 N）。
    /// 
    /// 注意 ReadElementContentAsString 的位置语义：它消费完 &lt;t&gt;...&lt;/t&gt; 后会把
    /// 读取器停在 **该元素之后的下一个节点**（通常是父级的 EndElement）。因此判定
    /// &lt;si&gt; 结束必须在消费子元素之前检查当前节点，否则会读到下一个 &lt;si&gt; 去，
    /// 导致所有条目被拼成一条（实测 360 条会退化成 1 条）。
    /// </summary>
    internal static List<string> LoadSharedStrings(ZipArchive zip)
    {
        var list = new List<string>();
        if (zip.OpenPartReader("xl/sharedStrings.xml", XlsxXml.BufferMedium) is not { } reader) return list;
        using (reader)

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "si") continue;

            if (reader.IsEmptyElement) { list.Add(""); continue; }

            int siDepth = reader.Depth;
            string text = ReadSharedStringItem(reader, siDepth);
            list.Add(text);
        }

        return list;
    }

    /// <summary>
    /// 读取单个 &lt;si&gt; 的全部 &lt;t&gt; 文本（支持富文本 &lt;r&gt;&lt;t&gt; 分片）。
    /// 返回时读取器停在 &lt;/si&gt; 上或其后，调用方的外层循环继续推进。
    /// </summary>
    private static string ReadSharedStringItem(XmlReader reader, int siDepth)
    {
        var sb = new StringBuilder();

        // 当前位于 <si>。逐个读取子节点，遇到 </si> 收尾。
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == siDepth && reader.LocalName == "si")
                break;

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t" && !reader.IsEmptyElement)
            {
                // 消费后读取器停在 <t> 之后；下一轮循环的 Read() 继续推进
                sb.Append(reader.ReadElementContentAsString());

                // ReadElementContentAsString 可能正好停在 </si>，此时需在此判定收尾
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == siDepth)
                    break;
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// styles.xml → 日期样式索引集合。
    /// 链路：cellXfs[i] → numFmtId → 内置日期集合 或 numFmts 自定义 formatCode。
    /// </summary>
    internal static HashSet<int> LoadDateStyles(ZipArchive zip)
    {
        var result = new HashSet<int>();

        var customDateFormats = new HashSet<int>();
        var xfNumFmtIds = new List<int>();

        if (zip.OpenPartReader("xl/styles.xml", XlsxXml.BufferSmall) is { } reader)
        using (reader)
        {
            bool inCellXfs = false;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "numFmt":
                            {
                                int id = ParseInt(reader, "numFmtId", -1);
                                if (id >= 0 && IsDateFormatCode(reader.GetAttribute("formatCode")))
                                    customDateFormats.Add(id);
                                break;
                            }
                        case "cellXfs":
                            inCellXfs = true;
                            break;
                        case "xf":
                            if (inCellXfs) xfNumFmtIds.Add(ParseInt(reader, "numFmtId", 0));
                            break;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "cellXfs")
                {
                    inCellXfs = false;
                }
            }
        }

        for (int i = 0; i < xfNumFmtIds.Count; i++)
        {
            int fmt = xfNumFmtIds[i];
            if (BuiltInDateFormats.Contains(fmt) || customDateFormats.Contains(fmt))
                result.Add(i);
        }

        return result;
    }

    private static int ParseInt(XmlReader reader, string name, int fallback)
    {
        string? v = reader.GetAttribute(name);
        return v != null && int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out int r) ? r : fallback;
    }

    /// <summary>
    /// 值是否为「数字型」（列级对齐的判定单元）。数字列右对齐是 Excel 的通用惯例。
    /// 
    /// 只认真正的数（含千分位、正负号、百分号、货币前缀）与日期；
    /// 带中文/字母的混合串（如「A12」「12个」）不算 —— 否则编号类文本列会被误判成数字列。
    /// </summary>
    private static bool IsNumericValue(object? value)
    {
        if (value is null) return false;
        if (value is double or int or long or decimal or float) return true;
        if (value is DateTime) return true;
        if (value is not string s || s.Length == 0) return false;

        return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out _);
    }

    /// <summary>formatCode 是否含日期/时间占位符（剥离引号字面量与转义后判定）</summary>
    private static bool IsDateFormatCode(string? formatCode)
    {
        if (string.IsNullOrEmpty(formatCode)) return false;

        bool inQuote = false, escaped = false;
        foreach (char c in formatCode)
        {
            if (escaped) { escaped = false; continue; }
            if (c == '\\') { escaped = true; continue; }
            if (c == '"') { inQuote = !inQuote; continue; }
            if (inQuote) continue;
            if (c is 'y' or 'Y' or 'm' or 'M' or 'd' or 'D') return true;
        }
        return false;
    }
}
