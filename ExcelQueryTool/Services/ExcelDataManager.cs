using System.Diagnostics;
using System.IO;
using System.Threading.Channels;

namespace ExcelQueryTool.Services
{
    /// <summary>
    /// 工作表数据加载器。
    /// 
    /// 数据源为自写的 <see cref="XlsxSheetReader"/>（性能说明见该类注释），
    /// 不再经过 MiniExcel 的 ExpandoObject 行映射。
    /// 
    /// 加载为<b>单趟扫描 + 真流式</b>：
    /// 扫描线程按有界 <see cref="Channel{T}"/> 逐行投递，消费侧立即 yield 出去，
    /// 因此内存占用为「通道容量 × 行宽」而非「全部命中行」。
    /// 
    /// 行为约定：
    /// - 表头行自动下探（第 1 行是整表大标题时取第 2 行，规则见 <see cref="XlsxSheetReader"/>），
    ///   数据从表头行之后开始
    /// - 日期由单元格 numFmt 判定后格式化为 yyyy-MM-dd
    /// - 空表头列一律保留并命名 &lt;列N&gt;
    /// - 表头重名追加 _2 / _3
    /// </summary>
    public class ExcelDataManager(ImageManager _imageManager)
    {
        /// <summary>
        /// 列占用探测的行数上限。空列判定只需抽样：前若干行已出现非空值的列即视为「有数据」。
        /// 有界探测保证结构判定不随数据量增长（1 万行以内即可稳定判定）。
        /// </summary>
        private const int OccupancyProbeRows = 10_000;

        /// <summary>
        /// 行通道容量（有界）。扫描侧远快于 UI 侧消费速度，容量起到背压作用：
        /// 消费跟得上时内存水位就是「容量 × 行宽」，不会随总行数增长。
        /// </summary>
        private const int RowChannelCapacity = 8192;

        /// <summary>
        /// 按文件路径缓存的包会话。
        ///
        /// 为什么放在这一层：会话要跨「列元数据 → 数据加载 → 切表 → 搜索 → 勾图片重载」
        /// 长期存活才能省掉重复开包，而这几次调用各自 <c>new XlsxSheetReader</c>，
        /// 会话挂在 reader 上等于每次新建即丢弃（实测省不下任何成本）。
        /// 放在 manager 上（随窗口存活、换文件时替换）才真正复用。
        ///
        /// 只保留最近一个文件：换文件后旧会话立即释放，不会累积文件句柄。
        /// </summary>
        private XlsxPackageSession? _session;
        private string? _sessionFilePath;

        /// <summary>取得（必要时创建）指定文件的包会话；换文件时释放旧的</summary>
        private XlsxPackageSession GetSession(string filePath)
        {
            if (_session != null && string.Equals(_sessionFilePath, filePath, StringComparison.Ordinal))
                return _session;

            _session?.Dispose();
            _session = new XlsxPackageSession(filePath);
            _sessionFilePath = filePath;
            return _session;
        }

        /// <summary>释放当前文件的包会话（换文件 / 关窗时调用）</summary>
        public void ReleaseSession()
        {
            _session?.Dispose();
            _session = null;
            _sessionFilePath = null;
        }

        /// <summary>
        /// 单趟流式加载：一次扫描同时完成结构判定与数据产出。
        /// 
        /// 表头到达时先做有界占用抽样（<see cref="OccupancyProbeRows"/> 行），
        /// 抽样结束即确定列映射，之后每一行立即按映射产出并投递到通道 ——
        /// 不再有「先扫完整表再产出」的全量缓冲，也不再需要开两趟解析同一个文件。
        /// </summary>
        /// <param name="expectedDisplayColumnCount">
        /// 调用方已按列元数据建好的显示列数（不含行号列）。传入后本方法会校验实际列数是否一致：
        /// 不一致说明两次独立扫描之间文件被改写，直接以异常终止，避免按过时映射静默错位。
        /// 传 null 表示跳过校验。
        /// </param>
        /// <param name="knownLayout">
        /// 调用方已取得的表结构（通常来自 <see cref="GetColumnMetadataWithLayout"/> 内部的那次扫描）。
        ///
        /// 传入后 <see cref="XlsxSheetReader.Scan"/> 会跳过重复开包与部件解析 ——
        /// UI 流程本就要先读一次包拿列元数据，再读一次包取数据行，每次开包都要重建
        /// ZipArchive 并重读 sharedStrings / styles。复用的安全性由「两条路径共用同一套
        /// 保守列映射（显示列序 == 源列号，见下方加载注释）」保证。
        ///
        /// 传 null 时退化为自行开包，行为与传入前完全一致。
        /// </param>
        /// <param name="prepared">
        /// 关键词解析结果（调用方 <see cref="SearchService.Prepare"/> + <see cref="SearchService.BindColumns"/>
        /// 一次性完成，树与词表、列限定绑定、高亮词表由窗口层共享）；null = 无过滤（整表浏览）。
        /// </param>
        public async IAsyncEnumerable<object[]> LoadWorksheetDataAsync(
            string filePath,
            string worksheetName,
            SearchService.PreparedSearch? prepared,
            bool showImages,
            [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken token,
            IProgress<string>? progress = null,
            int? expectedDisplayColumnCount = null,
            SheetLayout? knownLayout = null,
            TextTableReader.TextLayout? knownTextLayout = null)
        {
            token.ThrowIfCancellationRequested();

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                progress?.Report("文件不存在");
                yield break;
            }

            progress?.Report("正在分析工作表结构...");

            // 文本表（CSV/TXT）：解析走 TextTableReader，过滤/背压与 xlsx 路径共用
            if (TextTableReader.IsTextFile(filePath))
            {
                // 文本路径没有独立的「结构分析」阶段（一趟边判定边读），直接报加载数据
                progress?.Report("正在加载数据...");
                await foreach (var rowData in LoadTextRowsAsync(filePath, prepared, token, progress, expectedDisplayColumnCount, knownTextLayout)
                                   .ConfigureAwait(false))
                {
                    yield return rowData;
                }
                yield break;
            }

            // 有界通道背压（容量 RowChannelCapacity）；SingleReader 供 await foreach 独占消费
            var channel = CreateRowChannel();

            var reader = new XlsxSheetReader(filePath).UseSession(GetSession(filePath));
            int? expectedColumnCount = expectedDisplayColumnCount;
            string? layoutMismatch = null;

            // 扫描线程：表头行到达时确定列映射（含首次的有界占用抽样），之后逐行投递
            var scanTask = Task.Run(async () =>
            {
                // 列映射恒等（显示列序 == 源列号，下标 0 = RowNumber），只需记列数
                int columnCount = 0;
                HashSet<int>? imageColumns = null;
                long emitted = 0;
                int headerRow = 1;

                try
                {
                    reader.Scan(
                        worksheetName,
                        token,
                        OccupancyProbeRows,
                        // 表头到达：确定列映射。返回 false 表示空表，无需继续扫数据行
                        onLayout: layout =>
                        {
                            if (layout.ColumnCount == 0) return false;

                            headerRow = layout.HeaderRow > 0 ? layout.HeaderRow : 1;

                            // 保守列映射：显示列序 == 源列号（恒等）。
                            // 空表头列一律保留（不按占用剔除）：数据加载路径必须在首行数据前
                            // 确定列映射，此刻占用统计不可能完成；统一保守保留后空表头列固定
                            // 显示为 &lt;列N&gt;，行为可预测且绝不丢列（下标 0 = RowNumber 占位）。
                            columnCount = layout.ColumnCount;
                            // 图片锚点列可能落在数据区之外（如数据到 8 列、图锚在 10 列）：
                            // 只保留本表实际存在的列，行数组按列数分配，越界的列号会索引失败
                            var pictureCols = _imageManager.GetPictureColumns(worksheetName);
                            imageColumns = new HashSet<int>();
                            foreach (int col in pictureCols)
                                if (col <= columnCount) imageColumns.Add(col);

                            // 列结构守卫：与 UI 已建好的列数比对，两次独立扫描之间文件若被改写，
                            // 按旧映射取值会静默错位 —— 这里显式暴露（守卫消息与文本路径同一措辞）
                            if (expectedColumnCount is { } expected && expected != columnCount)
                            {
                                layoutMismatch = LayoutMismatchMessage("工作表", expected, columnCount);
                                return false;
                            }
                            // 结构分析到此结束、行加载即将开始：换文案。否则「正在分析工作表结构」
                            // 会替整个行加载阶段背锅（百万行该阶段 ~4-5s，结构分析本身仅 ~0.5s）
                            progress?.Report("正在加载数据...");
                            return true;
                        },
                        // 数据行：按映射取值、按需写入通道
                        onRow: (physicalRow, row) =>
                        {
                            var rowData = BuildRowData(row, physicalRow, columnCount, imageColumns!, showImages, worksheetName, headerRow);
                            if (!ShouldEmitRow(rowData, prepared)) return true;

                            // 此处 BuildRowData 已用物理行号装完图片，EmitRow 覆盖 rowData[0] 无副作用
                            EmitRow(channel.Writer, rowData, prepared, ref emitted, token);
                            return true;
                        },
                        // 复用调用方已有的布局：省掉一次完整的开包 + 部件解析
                        knownLayout: knownLayout,
                        // 把图片锚点列注入结构判定：图片常锚在数据区之外（如数据到第 8 列、
                        // 图片锚在第 10 列），列数不含这些列时「显示图片」永远渲染不出来
                        pictureColumnsProbe: _imageManager.GetPictureColumns,
                        // 头部截停：布局/占用/数字采样全部在头部完成，之后交并行尾巴。
                        // 快路径不可用时（XmlReader 回退）本参数无效，整表照旧扫完 —— 行为不变。
                        stopAfterDataRows: OccupancyProbeRows);

                    // 并行尾巴：剩余行切块多核解析+转换，按块序归并投递。
                    // tail==null = 快路径不可用（XmlReader 回退已扫完整表），无尾巴可扫。
                    var tail = reader.LastTailContext;
                    if (tail != null)
                    {
                        try
                        {
                            EmitTailInParallel(tail);
                        }
                        finally
                        {
                            tail.Release();
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    // 取消：下游通过 token 感知，此处只需收尾通道
                }
                catch (Exception ex)
                {
                    channel.Writer.TryComplete(ex);
                    return;
                }

                // 并行尾巴：按 <row 边界把尾区切成 K 块，worker 并行「解析+行转换」，
                // 本线程（merger）按块序过滤、投递 —— 保持行序与 Channel 单写者不变。
                // worker 侧不做图片解析（ImageManager 缓存非线程安全），图片列由 merger 补装；
                // sharedStrings / dateStyles 尾区期间只读，可并发。
                // 硬契约：头部已投行，任何块失败只能异常穿透，绝不回退重扫（会双投）。
                void EmitTailInParallel(ParallelTailContext tail)
                {
                    int workerCount = Math.Clamp(Environment.ProcessorCount - 1, 1, 8);
                    var starts = XlsxByteScanner.FindRowStarts(tail.Buffer, tail.TailStart, tail.Total);

                    // 块边界与顺序号种子：seed = 尾区前元素数 + 块起点之前的候选行数
                    // （顺序号语义 = 「第几个 row 元素」，缺 r 属性的行据此得到与全量扫描一致的行号）
                    var bounds = new List<int>(workerCount + 2) { tail.TailStart };
                    var seeds = new List<int>(workerCount + 2) { tail.RowsBeforeTail };
                    int pieces = Math.Min(workerCount, starts.Count + 1);
                    for (int c = 1; c < pieces; c++)
                    {
                        int idx = (int)((long)c * starts.Count / pieces);
                        bounds.Add(starts[idx]);
                        seeds.Add(tail.RowsBeforeTail + idx);
                    }
                    bounds.Add(tail.Total);

                    int chunkCount = bounds.Count - 1;
                    var tasks = new Task<List<object[]>>[chunkCount];
                    for (int c = 0; c < chunkCount; c++)
                    {
                        int start = bounds[c], end = bounds[c + 1], seed = seeds[c];
                        tasks[c] = Task.Run(() =>
                        {
                            var rows = new List<object[]>();
                            XlsxByteScanner.ScanChunk(
                                tail.Buffer.AsSpan(start, end - start), seed,
                                tail.SharedStrings, tail.DateStyles, token,
                                (physicalRow, row) =>
                                {
                                    rows.Add(BuildRowData(row, physicalRow, columnCount, imageColumns!, showImages, worksheetName, headerRow, resolveImages: false));
                                    return true;
                                });
                            return rows;
                        }, token);
                    }

                    try
                    {
                        for (int c = 0; c < chunkCount; c++)
                        {
                            var rows = tasks[c].GetAwaiter().GetResult();
                            foreach (var rowData in rows)
                            {
                                if (!ShouldEmitRow(rowData, prepared)) continue;

                                // 图片补装必须用原物理行号（rowData[0] 反推），重排要在其后
                                AttachImages(rowData, worksheetName, headerRow, imageColumns!, showImages);

                                // 搜索模式：行号重排为命中序号（与头部路径同语义）
                                EmitRow(channel.Writer, rowData, prepared, ref emitted, token);
                            }
                        }
                    }
                    catch
                    {
                        // 等其余 worker 收场，避免未观察异常；保留第一个异常向上抛
                        try { Task.WaitAll(tasks); } catch { /* 并发块的异常已由首个代表 */ }
                        throw;
                    }
                }

                channel.Writer.TryComplete();
            }, token);

            // 消费侧：逐行产出，扫描与 UI 渲染真正并行
            await foreach (var rowData in channel.Reader.ReadAllAsync(token).ConfigureAwait(false))
            {
                yield return rowData;
            }

            // 通道读完后再等扫描线程，确保异常/取消已被观察而非吞掉
            await scanTask.ConfigureAwait(false);

            // 列结构在两次扫描之间发生变化：显式失败，不返回错位数据
            if (layoutMismatch != null)
                throw new InvalidOperationException(layoutMismatch);
        }

        /// <summary>
        /// 文本表（CSV/TXT）数据行：布局已知（<paramref name="knownTextLayout"/> = prepare 阶段
        /// ReadLayout 的透传结果）时直接产出定宽行，省掉第二遍全量扫描；null 时回退自行 ReadLayout。
        /// 无图片列；关键词过滤、行号语义（第 1 条数据 = 1）、背压通道与 xlsx 路径共用。
        ///
        /// 数据段有两条路径，与 xlsx 的「头部 + 并行尾巴」同构：
        /// ① 字节级快路径（<see cref="TextTableReader.TextLayout.Fast"/> 非 null，即 UTF-8/GB18030）：
        ///    整文件字节已在 prepare 阶段读入，数据行切块多核「按区间解码 + 行转换」，
        ///    本线程（merger）按块序过滤、投递 —— 行序与 Channel 单写者不变，背压由合并端维持。
        /// ② char 级回退（BOM-less UTF-16 / 非字节级编码）：原 <see cref="TextTableReader.EnumerateRows"/>
        ///    单线程产出，行为与修改前完全一致。
        /// </summary>
        private async IAsyncEnumerable<object[]> LoadTextRowsAsync(
            string filePath,
            SearchService.PreparedSearch? prepared,
            [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken token,
            IProgress<string>? progress,
            int? expectedDisplayColumnCount,
            TextTableReader.TextLayout? knownTextLayout)
        {
            var channel = CreateRowChannel();

            var scanTask = Task.Run(() =>
            {
                long emitted = 0;
                try
                {
                    var layout = knownTextLayout ?? TextTableReader.ReadLayout(filePath);

                    // 列结构守卫：与 xlsx 路径同语义（守卫消息由 LayoutMismatchMessage 统一措辞）
                    if (expectedDisplayColumnCount is { } expected && expected != layout.ColumnCount)
                    {
                        channel.Writer.TryComplete(new InvalidOperationException(
                            LayoutMismatchMessage("文本文件", expected, layout.ColumnCount)));
                        return;
                    }

                    // 行号槽契约：各产出路径（xlsx BuildRowData / 文本 / EmitRow 搜索重排）一律 int——
                    // RowView.RowNumber 用 `is int` 模式匹配取值，long 会被判 0 致行头空白（v3.7 回归教训）
                    //
                    // 数据起始下标：HasHeader 时首行（下标 0）是表头，数据从下标 1 起；
                    // 否则数据从下标 0 起。与 EnumerateRows 的 headerPending 跳过口径一致。
                    int dataStart = layout.HasHeader ? 1 : 0;

                    if (layout.Fast is { } fast)
                        EmitTextRowsInParallel(fast, layout.ColumnCount, dataStart, prepared, channel.Writer, ref emitted, token);
                    else
                        EmitTextRowsChar(filePath, layout, prepared, channel.Writer, ref emitted, token);

                    channel.Writer.TryComplete();
                }
                catch (OperationCanceledException)
                {
                    // 取消：下游通过 token 感知，此处只需收尾通道
                    channel.Writer.TryComplete();
                }
                catch (Exception ex)
                {
                    channel.Writer.TryComplete(ex);
                    return;
                }
            }, token);

            await foreach (var rowData in channel.Reader.ReadAllAsync(token).ConfigureAwait(false))
            {
                yield return rowData;
            }

            // 通道读完后再等扫描线程，确保异常/取消已被观察而非吞掉
            await scanTask.ConfigureAwait(false);
        }

        /// <summary>
        /// 字节级数据段（UTF-8/GB18030 快路径）：把 <paramref name="fast"/> 的数据行区间切 K 块，
        /// worker 并行「按区间解码 + 行转换」，本线程按块序过滤、投递（与 xlsx <c>EmitTailInParallel</c> 同构）。
        /// worker 只解码字符串、不碰图片/ImageManager，无共享可变状态；背压由合并端 <see cref="EmitRow"/> 维持。
        /// 硬契约：worker 任何异常只能穿透，绝不回退重扫（会双投）。
        /// </summary>
        private static void EmitTextRowsInParallel(
            TextTableReader.TextFastData fast,
            int columnCount,
            int dataStart,
            SearchService.PreparedSearch? prepared,
            ChannelWriter<object[]> writer,
            ref long emitted,
            CancellationToken token)
        {
            int rowCount = fast.RowRanges.Length;
            int dataRows = rowCount - dataStart;
            if (dataRows <= 0) return;

            int workerCount = Math.Clamp(Environment.ProcessorCount - 1, 1, 8);
            int chunkCount = Math.Min(workerCount, dataRows);
            int baseLen = dataRows / chunkCount;
            int rem = dataRows % chunkCount;

            var tasks = new Task<List<object[]>>[chunkCount];
            for (int c = 0; c < chunkCount; c++)
            {
                int from = c * baseLen + Math.Min(c, rem);
                int len = baseLen + (c < rem ? 1 : 0);
                int startRange = dataStart + from;
                int endRange = startRange + len;

                tasks[c] = Task.Run(() =>
                {
                    // 整块一次解码（免逐行密集小解码），再逐行转定宽 + 行号占位
                    // 行号在本块内从 1 递增；合并端按块序重排为全局序号（见下方合并循环）
                    var fieldsRows = fast.DecodeRowsChunk(startRange, endRange, columnCount, token);
                    var rows = new List<object[]>(fieldsRows.Count);
                    for (int i = 0; i < fieldsRows.Count; i++)
                    {
                        string?[] fields = fieldsRows[i];
                        var rowData = new object[columnCount + 1];
                        for (int col = 0; col < columnCount; col++) rowData[col + 1] = fields[col] ?? "";
                        rows.Add(rowData);
                    }
                    return rows;
                }, token);
            }

            // 合并端：按块序（= 全局行序）逐行定序、定行号、过滤、投递 —— 行序与背压与 xlsx 完全一致
            int seq = 0;
            try
            {
                for (int c = 0; c < chunkCount; c++)
                {
                    var rows = tasks[c].GetAwaiter().GetResult();
                    foreach (var rowData in rows)
                    {
                        seq++;
                        rowData[0] = seq;
                        if (!ShouldEmitRow(rowData, prepared)) continue;

                        // 搜索模式：行号重排为命中序号（与 xlsx 路径同语义）
                        EmitRow(writer, rowData, prepared, ref emitted, token);
                    }
                }
            }
            catch
            {
                // 等其余 worker 收场，避免未观察异常；保留第一个异常向上抛
                try { Task.WaitAll(tasks); } catch { /* 并发块的异常已由首个代表 */ }
                throw;
            }
        }

        /// <summary>char 级数据段（回退路径）：原 <see cref="TextTableReader.EnumerateRows"/> 单线程产出。</summary>
        private static void EmitTextRowsChar(
            string filePath,
            TextTableReader.TextLayout layout,
            SearchService.PreparedSearch? prepared,
            ChannelWriter<object[]> writer,
            ref long emitted,
            CancellationToken token)
        {
            int seq = 0;
            foreach (var fields in TextTableReader.EnumerateRows(filePath, layout, token))
            {
                token.ThrowIfCancellationRequested();

                seq++;
                var rowData = new object[layout.ColumnCount + 1];
                rowData[0] = seq;
                for (int c = 0; c < layout.ColumnCount; c++) rowData[c + 1] = fields[c] ?? "";

                if (!ShouldEmitRow(rowData, prepared)) continue;

                // 搜索模式：行号重排为命中序号（与 xlsx 路径同语义）
                EmitRow(writer, rowData, prepared, ref emitted, token);
            }
        }

        /// <summary>
        /// 创建行产出通道：有界（容量 <see cref="RowChannelCapacity"/>）+ 背压 + 单写者单读者。
        /// 扫描侧远快于 UI 侧消费速度，容量决定内存水位「容量 × 行宽」，不随总行数增长。
        /// xlsx / 并行尾巴 / 文本三条路径共用。
        /// </summary>
        private static Channel<object[]> CreateRowChannel()
            => Channel.CreateBounded<object[]>(new BoundedChannelOptions(RowChannelCapacity)
            {
                SingleReader = true,
                SingleWriter = true,
                FullMode = BoundedChannelFullMode.Wait,
            });

        /// <summary>
        /// 行是否应投递：搜索模式走关键词匹配，浏览模式保留「至少一格有数据」的行。
        /// 三条产出路径（xlsx 头部 / 并行尾巴 / 文本）共用同一判定。
        /// </summary>
        private static bool ShouldEmitRow(object[] rowData, SearchService.PreparedSearch? prepared)
            => prepared != null ? SearchService.Matches(prepared, rowData) : HasAnyData(rowData);

        /// <summary>
        /// 计数 + 行号重排 + 写通道（同步阻塞写，背压即等待）。
        /// 搜索模式行号重排为命中序号（1,2,3…连续）；浏览模式保留原表位置（不覆盖 rowData[0]）。
        /// 三条产出路径共用；调用前若需图片补装必须先用原物理行号完成（重排在补装之后）。
        /// </summary>
        private static void EmitRow(
            ChannelWriter<object[]> writer,
            object[] rowData,
            SearchService.PreparedSearch? prepared,
            ref long emitted,
            CancellationToken token)
        {
            emitted++;
            if (prepared != null) rowData[0] = (int)emitted;
            writer.WriteAsync(rowData, token).AsTask().GetAwaiter().GetResult();
        }

        /// <summary>
        /// 列结构守卫的报错文案：两次独立扫描之间文件被改写时显式失败，不按过时映射静默错位。
        /// xlsx / 文本两条路径同一措辞，便于排查。
        /// </summary>
        private static string LayoutMismatchMessage(string subject, int expected, int actual)
            => $"{subject}列结构已变化（表头 {expected} 列，数据 {actual} 列），请重新打开文件";

        private object[] BuildRowData(
            XlsxRowBuffer row,
            int physicalRow,
            int columnCount,
            HashSet<int> imageColumns,
            bool showImages,
            string worksheetName,
            int headerRow,
            bool resolveImages = true)
        {
            var rowData = new object[columnCount + 1];
            // RowNumber：显示为「数据区内的序号」（第 1 条数据 = 1）。
            // 不再直接用 physicalRow - 1：表头行可能下探到第 2/3 行，物理行号与序号会差开，
            // 直接用物理行号会让首条数据显示为 2 或 3。
            // 搜索模式下过滤方会把命中行覆盖为连续命中序号（见 LoadWorksheetDataAsync 的 emitted 重排）。
            rowData[0] = physicalRow - headerRow;

            for (int displayCol = 1; displayCol <= columnCount; displayCol++)
            {
                // 恒等映射：显示列序即源列号
                int sourceCol = displayCol;
                object? cellValue = row.GetByColumn(sourceCol);
                var text = XlsxCellValue.ToDisplayString(cellValue);

                // resolveImages=false：并行 worker 侧只产文本（ImageManager 非线程安全），
                // 图片列由 merger 侧 AttachImages 补装
                if (showImages && resolveImages && imageColumns.Contains(displayCol))
                {
                    // LoadCellImages 返回的是缓存持有的实例（同格多图按锚点顺序）；
                    // 缓存已保证尺寸有效，单元格只保存引用，不负责释放（详见 ImageTextPair 注释）
                    var imgs = _imageManager.LoadCellImages(worksheetName, physicalRow, sourceCol);
                    rowData[displayCol] = imgs.Count > 0
                        ? new ImageTextPair { Images = imgs, Text = text }
                        : text;
                }
                else
                {
                    rowData[displayCol] = text;
                }
            }

            return rowData;
        }

        /// <summary>
        /// 并行尾巴产出的行：图片列在 merger 侧（单线程）补装 ImageTextPair。
        /// 行还在过滤前不装图片，被过滤掉的行连图片加载都省掉。
        /// </summary>
        private void AttachImages(object[] rowData, string worksheetName, int headerRow, HashSet<int> imageColumns, bool showImages)
        {
            if (!showImages) return;

            int physicalRow = (int)rowData[0] + headerRow;
            foreach (int displayCol in imageColumns)
            {
                if (rowData[displayCol] is string text)
                {
                    var imgs = _imageManager.LoadCellImages(worksheetName, physicalRow, displayCol);
                    if (imgs.Count > 0)
                        rowData[displayCol] = new ImageTextPair { Images = imgs, Text = text };
                }
            }
        }

        private static bool HasAnyData(object[] rowData)
        {
            for (int i = 1; i < rowData.Length; i++)
            {
                if (rowData[i] is string s && s.Length > 0) return true;
                if (rowData[i] is ImageTextPair) return true;
            }
            return false;
        }



        /// <summary>
        /// 工作表名列表。文本表（CSV/TXT）只有一个伪工作表：文件名（不含扩展名）。
        /// </summary>
        public List<string> GetWorksheetNames(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return [];
            if (TextTableReader.IsTextFile(filePath)) return [Path.GetFileNameWithoutExtension(filePath)];
            // 复用包会话：工作表名列表与随后的列元数据 / 数据加载共用同一个已打开的包
            return XlsxSheetReader.GetSheetNames(GetSession(filePath));
        }

        /// <summary>
        /// 列元数据 + 产生它的表结构。
        ///
        /// 表头 + 图片列标记；空表头列一律保留并命名 &lt;列N&gt;，重名追加 _2/_3。
        /// 之所以把布局一并交出去：本方法内部必然做了完整一次「开包 + 表头 + 占用抽样」，
        /// 而列映射规则（保守保留空表头列）已使该布局与数据加载路径所需的完全等价，
        /// 调用方把它透传给 <see cref="LoadWorksheetDataAsync"/> 即可省掉第二次开包。
        ///
        /// 实例方法（而非 static）：要复用本 manager 的包会话，让这次扫描与随后的数据加载
        /// 共用同一个已打开的包，而不是各开一次（实测开包占单表成本 61%）。
        /// </summary>
        public (List<ColumnMeta> Columns, SheetLayout? Layout, TextTableReader.TextLayout? TextLayout) GetColumnMetadataWithLayout(
            string filePath, string worksheetName, ImageManager imageManager)
        {
            List<ColumnMeta> columns = [];
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return (columns, null, null);

            // 文本表（CSV/TXT）：结构由 TextTableReader 全量扫描得出，无 SheetLayout 可复用；
            // TextLayout 原样透传给数据加载，省掉数据段的第二遍全量 ReadLayout。列命名与 xlsx 侧同一套规则
            if (TextTableReader.IsTextFile(filePath))
            {
                // 字节级快路径：一次字节扫描得列数+表头（首屏门控从 ~530ms char 全量扫描降到 ~57ms）。
                // UTF-8/GB18030 走快路径；BOM-less UTF-16 / 非字节级编码回退 char 级 ReadLayout，行为不变。
                var textLayout = TextTableReader.TryReadFast(filePath, out var fast) ? fast
                    : TextTableReader.ReadLayout(filePath);
                return (TextTableReader.BuildColumns(textLayout), null, textLayout);
            }

            SheetLayout layout;
            try
            {
                // 与数据加载同源：同一个有界抽样探测，保证列集合与顺序完全一致。
                // 图片锚点列必须同样注入 —— 两条路径的列数若不一致（元数据含图片列、数据不含，
                // 或反之），表头与行数据会整体错位，这是本项目最硬的一条约束。
                layout = new XlsxSheetReader(filePath).UseSession(GetSession(filePath)).ReadLayoutWithOccupancy(
                    worksheetName,
                    CancellationToken.None,
                    OccupancyProbeRows,
                    imageManager.GetPictureColumns);
            }
            catch (Exception ex)
            {
                // 结构探测失败不阻塞：列元数据照常返回，仅无表结构信息
                Debug.WriteLine($"读取工作表结构失败: {ex.Message}");
                return (columns, null, null);
            }

            if (layout.ColumnCount == 0) return (columns, null, null);

            var pictureColumns = imageManager.GetPictureColumns(worksheetName);
            var normalizer = new HeaderNormalizer();

            // 与数据加载共用同一套列映射（恒等：显示列序 == 源列号），确保表头与行数据不会错位
            for (int sourceCol = 1; sourceCol <= layout.ColumnCount; sourceCol++)
            {
                string header = layout.Headers.Length > sourceCol ? layout.Headers[sourceCol] ?? "" : "";
                string name = normalizer.Normalize(sourceCol, header);

                // Excel 原始列宽按源列取（未声明的列 ExcelWidth=0，UI 用默认宽度）
                double excelWidth = layout.ColumnWidths.Length > sourceCol ? layout.ColumnWidths[sourceCol] : 0;

                // 列级数字标记（数据区抽样得出）：供 UI 做「数字右对齐、文本左对齐」
                bool isNumeric = layout.ColumnIsNumeric.Length > sourceCol && layout.ColumnIsNumeric[sourceCol];

                columns.Add(new ColumnMeta(name, pictureColumns.Contains(sourceCol), excelWidth, isNumeric));
            }

            return (columns, layout, null);
        }

        /// <summary>
        /// 准备图片索引与列元数据，并确保工作表存在（自窗口 code-behind 层下沉）。
        ///
        /// 图片索引<b>按需构建</b>：只在勾选「显示图片」且该文件尚未建过索引时才扫包 ——
        /// 建索引是打开流程里最贵的一步（109 MB 的包实测 637 ms，占整个打开段九成），
        /// 不看图片时完全跳过，文件打开立即提速。勾选「显示图片」会触发整表重载并走到这里，
        /// 索引此时才补建，且按文件记忆（<see cref="ImageManager.HasPictureIndexFor"/>），
        /// 同一文件内切工作表不会重扫。
        ///
        /// 位图缓存则在每次进入时清掉：上一次的位图可能很大，且新工作表未必复用它们。
        /// 索引（轻量的 byte[] 引用）与位图缓存（昂贵的已解码位图）生命周期本就不同。
        /// </summary>
        public async Task<(List<ColumnMeta> Columns, bool Ok, SheetLayout? Layout, TextTableReader.TextLayout? TextLayout)> PrepareWorksheetAsync(
            string filePath, string worksheetName, bool showImages, CancellationToken token)
        {
            // 先回收上一次解码的位图；索引保留，下面按需判断是否要重建
            _imageManager.ClearImageCache();

            // 文本表（CSV/TXT）没有图片概念，也不可当 zip 扫描，直接跳过索引构建
            if (showImages
                && !TextTableReader.IsTextFile(filePath)
                && !_imageManager.HasPictureIndexFor(filePath))
            {
                // 图片索引需要可 Seek 的流，独立打开避免文件句柄冲突
                await Task.Run(() =>
                {
                    using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    _imageManager.BuildPictureIndex(stream, token, filePath);
                }, token);
            }

            // 一并取回表结构：数据加载阶段透传它即可省掉第二次开包 / 第二遍全量 ReadLayout（文本表）
            var (columns, layout, textLayout) = await Task.Run(
                () => GetColumnMetadataWithLayout(filePath, worksheetName, _imageManager), token);

            return (columns, columns.Count > 0, layout, textLayout);
        }
    }
}
