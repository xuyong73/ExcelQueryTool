using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace ExcelQueryTool.Services
{
    /// <summary>
    /// 直接解析 xlsx（OOXML）包，建立「工作表名 + 行号 + 列号 → 图片二进制」的索引。
    ///
    /// 这是 MiniExcel 的补齐层：MiniExcel 只处理单元格值，不读 drawing/media，
    /// 因此图片锚点信息由本类自行解析。
    ///
    /// 解析链路：
    ///   xl/workbook.xml           → sheet 名与 r:id 的映射
    ///   xl/_rels/workbook.xml.rels → r:id → worksheets/sheetN.xml
    ///   xl/worksheets/sheetN.xml  → &lt;drawing r:id="..."/&gt; + 列宽/行高几何 + vm 内嵌图单元格
    ///   xl/worksheets/_rels/sheetN.xml.rels → drawing 部件路径
    ///   xl/drawings/drawingN.xml  → anchor 与 blip r:embed
    ///   xl/drawings/_rels/drawingN.xml.rels → r:embed → media 部件路径
    ///   xl/media/*                → 图片二进制
    ///
    /// 图片归属格按「几何中心」定位：twoCellAnchor 取 from/to 中点、oneCellAnchor 取
    /// from + ext/2、absoluteAnchor 取 pos + ext/2——中心落在哪一格就归哪一格
    /// （列向按列宽加权，行向按行高加权，未声明的行列用默认值）。几何不可用时
    /// 回退到 from 左上角锚点（旧行为）。行列号统一转换为 1 基，与 ImageManager
    /// 的 (name, row, col) 约定一致。包内路径解析复用 <see cref="XlsxPartPath"/>。
    /// </summary>
    public sealed class XlsxPictureIndex
    {

        /// <summary>(工作表名, 行号 1 基, 列号 1 基) → 图片二进制列表（同格多图按包内锚点顺序）</summary>
        public Dictionary<(string Sheet, int Row, int Col), List<byte[]>> Pictures { get; } = [];

        /// <summary>
        /// 同一单元格可被多个锚点命中（一 cell 多图，如 E8 里贴了 4 张截图）：
        /// 追加而非覆盖 —— 曾经的单值字典会静默丢掉先到的图（实测 67 锚点只剩 34 张）。
        /// </summary>
        private void AddPicture(string sheetName, int row, int col, byte[] bytes)
        {
            if (bytes.Length == 0) return;
            var key = (sheetName, row, col);
            if (!Pictures.TryGetValue(key, out var bucket))
                Pictures[key] = bucket = [];
            bucket.Add(bytes);
        }

        /// <summary>
        /// 从 xlsx 流构建图片索引。流需可 Seek；调用方保留流的生命周期。
        /// </summary>
        public static XlsxPictureIndex Build(Stream stream)
        {
            var index = new XlsxPictureIndex();
            if (stream == null || !stream.CanRead) return index;

            try
            {
                using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

                // 1. sheet 名 → sheet 部件路径
                var sheetParts = ResolveSheetParts(zip);
                if (sheetParts.Count == 0) return index;

                // 缓存已解析的 media 部件，同一图片可被多个 anchor 引用
                var mediaCache = new Dictionary<string, byte[]>(StringComparer.Ordinal);

                foreach (var (sheetName, sheetPart) in sheetParts)
                {
                    // 一次流式扫描：中心定位几何（列宽/行高）+ vm 内嵌图单元格
                    var (geometry, vmCells) = ReadSheetGeometry(zip, sheetPart);

                    // 浮动/锚定图片：drawing 部件
                    var drawingPart = ResolveDrawingPart(zip, sheetPart);
                    if (drawingPart != null)
                    {
                        var imagesInPart = new Dictionary<string, byte[]>(StringComparer.Ordinal);
                        var anchors = ReadAnchors(zip, drawingPart, geometry, imagesInPart, mediaCache);

                        foreach (var (row, col, relId) in anchors)
                        {
                            if (imagesInPart.TryGetValue(relId, out var bytes))
                            {
                                index.AddPicture(sheetName, row, col, bytes);
                            }
                        }
                    }

                    // Excel 新版「单元格内嵌图片」：vm= 元数据 → richData → richValueRel → media。
                    // 这类图不在 drawing 部件里，缺这条链路时单元格只显示 #VALUE!。
                    // 内嵌图天然属于持有它的单元格，无需中心换算。
                    foreach (var (row, col, bytes) in ReadRichValueImages(zip, vmCells, mediaCache))
                    {
                        index.AddPicture(sheetName, row, col, bytes);
                    }
                }
            }
            catch
            {
                // 解析失败不阻塞数据加载 —— 图片退化为无图
            }

            return index;
        }

        /// <summary>读取 workbook.xml + 其 rels，得到 工作表名 → 部件路径（形如 xl/worksheets/sheet1.xml）</summary>
        private static List<(string Name, string Part)> ResolveSheetParts(ZipArchive zip)
        {
            var result = new List<(string, string)>();
            var rels = XlsxRelationships.ReadRels(zip, "xl/workbook.xml");

            foreach (var (name, rid) in XlsxRelationships.EnumerateSheets(zip))
            {
                if (!rels.TryGetValue(rid, out var rel) || rel.Type != XlsxRelationships.RelTypeWorksheet) continue;

                var part = XlsxPartPath.ResolvePartPath("xl/workbook.xml", rel.Target);
                if (part != null) result.Add((name, part));
            }

            return result;
        }

        /// <summary>从 sheetN.xml.rels 找到 drawing 部件路径</summary>
        private static string? ResolveDrawingPart(ZipArchive zip, string sheetPart)
        {
            var rels = XlsxRelationships.ReadRels(zip, sheetPart);

            foreach (var rel in rels.Values)
            {
                if (rel.Type != XlsxRelationships.RelTypeDrawing) continue;
                return XlsxPartPath.ResolvePartPath(sheetPart, rel.Target);
            }

            return null;
        }

        /// <summary>解析 drawingN.xml，返回 (行 1 基, 列 1 基, r:embed) 列表（中心定位），并填充图片二进制</summary>
        private static List<(int Row, int Col, string RelId)> ReadAnchors(
            ZipArchive zip,
            string drawingPart,
            SheetGeometry? geometry,
            Dictionary<string, byte[]> imagesOut,
            Dictionary<string, byte[]> mediaCache)
        {
            var anchors = new List<(int, int, string)>();

            var entry = zip.GetEntry(drawingPart);
            if (entry == null) return anchors;

            var doc = LoadXml(entry);
            if (doc == null) return anchors;

            var anchorNodes = doc.DocumentElement?.SelectNodes(
                "*[local-name()='twoCellAnchor' or local-name()='oneCellAnchor' or local-name()='absoluteAnchor']");

            if (anchorNodes == null) return anchors;

            foreach (XmlNode anchor in anchorNodes)
            {
                if (anchor.NodeType != XmlNodeType.Element) continue;

                // 图片的 blip embed 关系 id
                var blip = anchor.SelectSingleNode(".//*[local-name()='blip']");
                string? embed = null;
                if (blip?.Attributes != null)
                {
                    foreach (XmlAttribute attr in blip.Attributes)
                    {
                        if (attr.LocalName == "embed" && attr.NamespaceURI == XlsxRelationships.NsRelationships)
                        {
                            embed = attr.Value;
                            break;
                        }
                    }
                }

                if (string.IsNullOrEmpty(embed)) continue;

                (int row, int col) = ResolveAnchorCell(anchor, geometry);
                if (row < 0 || col < 0) continue;

                // 加载图片二进制（按 relId 缓存）
                if (!imagesOut.ContainsKey(embed))
                {
                    var bytes = LoadMedia(zip, drawingPart, embed, mediaCache);
                    if (bytes != null) imagesOut[embed] = bytes;
                }

                anchors.Add((row, col, embed));
            }

            return anchors;
        }

        /// <summary>
        /// 按图片几何中心定位归属单元格（返回 1 基行列）。三种锚定方式分别取中心：
        ///   twoCellAnchor  = (from, to) 的中点；
        ///   oneCellAnchor  = from + ext/2；
        ///   absoluteAnchor = pos + ext/2（旧版直接跳过不显示，现一并支持）。
        /// 几何不可用或换算失败时回退到 from 左上角（与旧行为一致）。
        /// </summary>
        private static (int Row, int Col) ResolveAnchorCell(XmlNode anchor, SheetGeometry? geo)
        {
            if (geo == null) return FromAnchorCell(anchor);

            try
            {
                string localName = anchor.LocalName ?? "";

                if (localName == "twoCellAnchor")
                {
                    var from = anchor.SelectSingleNode("*[local-name()='from']");
                    var to = anchor.SelectSingleNode("*[local-name()='to']");
                    if (from == null || to == null) return FromAnchorCell(anchor);

                    int fc = ReadIntChild(from, "col") ?? -1, fr = ReadIntChild(from, "row") ?? -1;
                    int tc = ReadIntChild(to, "col") ?? -1, tr = ReadIntChild(to, "row") ?? -1;
                    if (fc < 0 || fr < 0 || tc < 0 || tr < 0) return FromAnchorCell(anchor);

                    double fromPx = geo.ColCumPx(fc) + ReadEmuChild(from, "colOff") / SheetGeometry.EmuPerPx;
                    double toPx = geo.ColCumPx(tc) + ReadEmuChild(to, "colOff") / SheetGeometry.EmuPerPx;
                    double fromPt = geo.RowCumPt(fr) + ReadEmuChild(from, "rowOff") / SheetGeometry.EmuPerPoint;
                    double toPt = geo.RowCumPt(tr) + ReadEmuChild(to, "rowOff") / SheetGeometry.EmuPerPoint;
                    return (geo.FindRow((fromPt + toPt) / 2) + 1, geo.FindCol((fromPx + toPx) / 2) + 1);
                }

                if (localName == "oneCellAnchor")
                {
                    var from = anchor.SelectSingleNode("*[local-name()='from']");
                    var ext = anchor.SelectSingleNode("*[local-name()='ext']");
                    if (from == null || ext == null) return FromAnchorCell(anchor);

                    int fc = ReadIntChild(from, "col") ?? -1, fr = ReadIntChild(from, "row") ?? -1;
                    if (fc < 0 || fr < 0) return FromAnchorCell(anchor);

                    double centerPx = geo.ColCumPx(fc) + ReadEmuChild(from, "colOff") / SheetGeometry.EmuPerPx
                                      + ReadEmuAttr(ext, "cx") / SheetGeometry.EmuPerPx / 2;
                    double centerPt = geo.RowCumPt(fr) + ReadEmuChild(from, "rowOff") / SheetGeometry.EmuPerPoint
                                      + ReadEmuAttr(ext, "cy") / SheetGeometry.EmuPerPoint / 2;
                    return (geo.FindRow(centerPt) + 1, geo.FindCol(centerPx) + 1);
                }

                if (localName == "absoluteAnchor")
                {
                    var pos = anchor.SelectSingleNode("*[local-name()='pos']");
                    var ext = anchor.SelectSingleNode("*[local-name()='ext']");
                    if (pos == null || ext == null) return (-1, -1);   // 与旧行为一致：无法定位则不显示

                    double centerPx = (ReadEmuAttr(pos, "x") + ReadEmuAttr(ext, "cx") / 2) / SheetGeometry.EmuPerPx;
                    double centerPt = (ReadEmuAttr(pos, "y") + ReadEmuAttr(ext, "cy") / 2) / SheetGeometry.EmuPerPoint;
                    return (geo.FindRow(centerPt) + 1, geo.FindCol(centerPx) + 1);
                }
            }
            catch
            {
                // 几何换算的任何意外：回退 from 定位，图片不丢
            }

            return FromAnchorCell(anchor);
        }

        /// <summary>from 左上角锚点（0 基 +1 → 1 基）；无 from 返回 (-1,-1) 表示无法定位</summary>
        private static (int Row, int Col) FromAnchorCell(XmlNode anchor)
        {
            var from = anchor.SelectSingleNode("*[local-name()='from']");
            if (from == null) return (-1, -1);
            int? col = ReadIntChild(from, "col");
            int? row = ReadIntChild(from, "row");
            if (col == null || row == null) return (-1, -1);
            return (row.Value + 1, col.Value + 1);
        }

        /// <summary>
        /// 中心定位所需的表几何：列宽（字符宽 × 假设像素系数）与行高（磅）。
        ///
        /// 行高是精确值（OOXML 直接给磅），行方向的中心落点无字体假设；
        /// 列宽 OOXML 给「字符数」，换算像素依赖默认字体数字宽（MDW）——这里取 7px
        /// （Calibri/等线 11 @96dpi 的常见值）。MDW 只影响跨列图片中心落点的相对权重，
        /// 误差远小于一个列宽的零头；零宽（隐藏）列自然不参与落点。未声明的行列用默认值。
        /// </summary>
        private sealed class SheetGeometry
        {
            public const double PxPerChar = 7.0;
            public const double EmuPerPx = 9525.0;
            public const double EmuPerPoint = 12700.0;
            internal const int MaxColumns = 16384;
            private const int MaxRows = 1048576;

            public double[] ColPx = [];               // 1 基，[0] 未用；未声明的列用 DefaultColPx
            public double DefaultColPx = 8.43 * PxPerChar;
            public readonly Dictionary<int, double> RowPt = new();   // 1 基行号 → 高度（磅）
            public double DefaultRowPt = 15.0;

            public double ColPxAt(int col1) => col1 >= 1 && col1 < ColPx.Length ? ColPx[col1] : DefaultColPx;
            public double RowPtAt(int row1) => RowPt.TryGetValue(row1, out var v) ? v : DefaultRowPt;

            /// <summary>0 基列号 col0 之前（不含）的累计像素宽</summary>
            public double ColCumPx(int col0)
            {
                double s = 0;
                for (int i = 1; i <= col0; i++) s += ColPxAt(i);
                return s;
            }

            /// <summary>0 基行号 row0 之前（不含）的累计磅高</summary>
            public double RowCumPt(int row0)
            {
                double s = 0;
                for (int i = 1; i <= row0; i++) s += RowPtAt(i);
                return s;
            }

            /// <summary>绝对像素坐标落在哪一列（返回 0 基列号；零宽列不参与落点）</summary>
            public int FindCol(double centerPx)
            {
                double cum = 0;
                for (int i = 0; i < MaxColumns; i++)
                {
                    cum += ColPxAt(i + 1);
                    if (centerPx < cum) return i;
                }
                return 0;
            }

            /// <summary>绝对磅坐标落在哪一行（返回 0 基行号）</summary>
            public int FindRow(double centerPt)
            {
                double cum = 0;
                for (int i = 0; i < MaxRows; i++)
                {
                    cum += RowPtAt(i + 1);
                    if (centerPt < cum) return i;
                }
                return 0;
            }
        }

        /// <summary>
        /// 流式扫一遍 sheet 部件，收集中心定位所需的几何（列宽/行高/默认值）与 vm 内嵌图单元格。
        /// 解析失败返回 (null, [])：锚点回退到 from 定位，vm 图放弃（与旧行为一致）。
        /// </summary>
        private static (SheetGeometry? Geometry, List<(int Vm, int Row, int Col)> VmCells) ReadSheetGeometry(
            ZipArchive zip, string sheetPart)
        {
            var geo = new SheetGeometry();
            var vmCells = new List<(int, int, int)>();
            var colWidths = new Dictionary<int, double>();

            var sheetEntry = zip.GetEntry(sheetPart);
            if (sheetEntry == null) return (null, vmCells);

            try
            {
                using var stream = sheetEntry.Open();
                using var reader = XlsxXml.CreatePartReader(stream, XlsxXml.BufferSheet);

                int currentRow = 0, currentCol = 0;
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element) continue;
                    switch (reader.LocalName)
                    {
                        case "sheetFormatPr":
                            if (TryParseInvariant(reader.GetAttribute("defaultColWidth"), out double dcw))
                                geo.DefaultColPx = dcw * SheetGeometry.PxPerChar;
                            if (TryParseInvariant(reader.GetAttribute("defaultRowHeight"), out double drh))
                                geo.DefaultRowPt = drh;
                            break;

                        case "col":
                            if (int.TryParse(reader.GetAttribute("min"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int min)
                                && int.TryParse(reader.GetAttribute("max"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int max)
                                && TryParseInvariant(reader.GetAttribute("width"), out double width)
                                && min >= 1 && max >= min)
                            {
                                for (int c = min; c <= Math.Min(max, SheetGeometry.MaxColumns); c++)
                                    colWidths[c] = width;
                            }
                            break;

                        case "row":
                            if (int.TryParse(reader.GetAttribute("r"), out int rn)) currentRow = rn;
                            if (currentRow > 0 && TryParseInvariant(reader.GetAttribute("ht"), out double ht))
                                geo.RowPt[currentRow] = ht;
                            break;

                        case "c":
                            // r 形如 "J5"（字母+数字），取字母段算 1 基列号（实现与两条扫描路径共用）
                            string? cellRef = reader.GetAttribute("r");
                            if (!string.IsNullOrEmpty(cellRef))
                                currentCol = XlsxValueParsing.ParseColumnLetters(cellRef);
                            string? vm = reader.GetAttribute("vm");
                            if (vm != null && int.TryParse(vm, out int vmIdx))
                                vmCells.Add((vmIdx, currentRow, currentCol));
                            break;
                    }
                }
            }
            catch
            {
                // 部件损坏：整体回退（几何 null → 锚点退回 from 定位；vm 图放弃）
                return (null, []);
            }

            // 稀疏列宽声明转数组
            int maxCol = 0;
            foreach (var c in colWidths.Keys) maxCol = Math.Max(maxCol, c);
            geo.ColPx = new double[maxCol + 1];
            foreach (var (c, w) in colWidths)
                geo.ColPx[c] = w * SheetGeometry.PxPerChar;

            return (geo, vmCells);
        }

        private static bool TryParseInvariant(string? s, out double value)
            => double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out value);

        /// <summary>
        /// 解析 Excel 新版「单元格内嵌图片」链路，返回 (行 1 基, 列 1 基, 图片字节) 列表。
        ///
        /// 链路：sheet.xml 里 &lt;c vm="N" t="e"&gt; 单元格 → metadata.xml 的 valueMetadata/futureMetadata
        /// 把 vm 索引映射到 richValue 序号 → xl/richData/rdrichvalue.xml 第 N 个 &lt;rv&gt; 的
        /// _rvRel 键 → xl/richData/_rels/richValueRel.xml.rels 的 image 关系 → media 部件。
        /// 实测文件里 rv 值序与 vm 索引一一对应（&lt;v&gt;0&lt;/v&gt; 即 rv[0]），故直接按下标取。
        /// vmCells 由 <see cref="ReadSheetGeometry"/> 的同一趟扫描收集。
        /// </summary>
        private static List<(int Row, int Col, byte[] Bytes)> ReadRichValueImages(
            ZipArchive zip, List<(int Vm, int Row, int Col)> vmCells, Dictionary<string, byte[]> mediaCache)
        {
            var result = new List<(int, int, byte[])>();
            if (vmCells.Count == 0) return result;

            // 1. rdrichvalue.xml：第 i 个 rv 的 _rvRel 值 → richValueRel 关系 id
            var relIdByRv = ReadRichValueRelIds(zip);
            if (relIdByRv.Count == 0) return result;

            // 2. 按 vm 索引取关系 id。关系文件是 richValueRel.xml（其 .rels 与它同目录），
            // 而非 rdrichvalue.xml —— rdrichvalue 只提供 rv 下标 → rId 的映射。
            var rels = XlsxRelationships.ReadRels(zip, "xl/richData/richValueRel.xml");
            foreach (var (vm, row, col) in vmCells)
            {
                // vm 属性是 1 基（vm="1" → rv[0]），rv 下标是 0 基
                if (!relIdByRv.TryGetValue(vm - 1, out var relId)) continue;
                if (!rels.TryGetValue(relId, out var rel) || rel.Type != XlsxRelationships.RelTypeImage) continue;

                var partPath = XlsxPartPath.ResolvePartPath("xl/richData/rdrichvalue.xml", rel.Target);
                if (partPath == null) continue;

                byte[]? bytes;
                lock (mediaCache)
                {
                    if (!mediaCache.TryGetValue(partPath, out bytes))
                    {
                        var mediaEntry = zip.GetEntry(partPath);
                        if (mediaEntry == null) continue;
                        bytes = ReadEntryBytes(mediaEntry);
                        mediaCache[partPath] = bytes;
                    }
                }

                if (bytes.Length > 0) result.Add((row, col, bytes));
            }

            return result;
        }

        /// <summary>
        /// 解析内嵌图片的间接映射，返回 vm 下标（0 基）→ richValueRel 关系 id。
        ///
        /// 实际结构（实测 Excel 2408+ 生成的文件）：
        ///   rdrichvalue.xml: &lt;rv&gt;&lt;v&gt;0&lt;/v&gt;&lt;v&gt;5&lt;/v&gt;&lt;/rv&gt; —— &lt;v&gt; 无 t 属性，
        ///   第一个 &lt;v&gt; 的数值是 richValueRel.xml 里 &lt;rel&gt; 元素的<b>位置下标</b>；
        ///   richValueRel.xml: &lt;rel r:id="rId1"/&gt; —— 按出现顺序数下标，取其 r:id。
        /// 单元格 vm="N"（1 基）→ rv[N-1] → rel[第一个v] → r:id。
        /// </summary>
        private static Dictionary<int, string> ReadRichValueRelIds(ZipArchive zip)
        {
            var result = new Dictionary<int, string>();

            // richValueRel.xml：按出现顺序收集 rel 的 r:id
            var relIdsByPos = new List<string>();
            var rve = zip.GetEntry("xl/richData/richValueRel.xml");
            if (rve != null)
            {
                var rvDoc = LoadXml(rve);
                if (rvDoc?.DocumentElement != null)
                {
                    foreach (XmlNode rel in rvDoc.DocumentElement.ChildNodes)
                    {
                        if (rel.NodeType != XmlNodeType.Element) continue;
                        var attrs = rel.Attributes;
                        if (attrs == null) continue;

                        foreach (XmlAttribute attr in attrs)
                        {
                            if (attr.LocalName == "id" && attr.NamespaceURI == XlsxRelationships.NsRelationships)
                            {
                                relIdsByPos.Add(attr.Value);
                                break;
                            }
                        }
                    }
                }
            }
            if (relIdsByPos.Count == 0) return result;

            // rdrichvalue.xml：第 i 个 rv → 第一个 v 的数值 → relIdsByPos 下标
            var entry = zip.GetEntry("xl/richData/rdrichvalue.xml");
            if (entry == null) return result;

            var doc = LoadXml(entry);
            if (doc?.DocumentElement == null) return result;

            int i = 0;
            foreach (XmlNode rv in doc.DocumentElement.ChildNodes)
            {
                if (rv.NodeType != XmlNodeType.Element) continue;

                foreach (XmlNode v in rv.ChildNodes)
                {
                    if (v.NodeType != XmlNodeType.Element) continue;
                    // 第一个数值子节点：richValueRel 的位置下标（t 属性无或非 _rvRel 时也按此解析）
                    if (int.TryParse(v.InnerText.Trim(), out int pos) && pos >= 0 && pos < relIdsByPos.Count)
                    {
                        result[i] = relIdsByPos[pos];
                    }
                    break;   // 只看第一个 <v>
                }
                i++;
            }

            return result;
        }

        private static int? ReadIntChild(XmlNode parent, string localName)
        {
            var node = parent.SelectSingleNode($"*[local-name()='{localName}']");
            if (node?.InnerText is { Length: > 0 } text && int.TryParse(text.Trim(), out int val))
                return val;
            return null;
        }

        /// <summary>读子元素的 EMU 数值（colOff/rowOff）；缺失或非法按 0</summary>
        private static double ReadEmuChild(XmlNode parent, string localName)
            => double.TryParse(
                parent.SelectSingleNode($"*[local-name()='{localName}']")?.InnerText?.Trim(),
                NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0;

        /// <summary>读元素属性（absoluteAnchor 的 pos x/y 与 ext cx/cy，EMU）；缺失或非法按 0</summary>
        private static double ReadEmuAttr(XmlNode node, string name)
            => double.TryParse(node.Attributes?[name]?.Value?.Trim(),
                NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0;

        /// <summary>经 drawing rels 把 r:embed 解析为 media 部件路径并读出字节</summary>
        private static byte[]? LoadMedia(
            ZipArchive zip,
            string drawingPart,
            string relId,
            Dictionary<string, byte[]> mediaCache)
        {
            var rels = XlsxRelationships.ReadRels(zip, drawingPart);
            if (!rels.TryGetValue(relId, out var rel)) return null;
            if (rel.Type != XlsxRelationships.RelTypeImage) return null;

            var partPath = XlsxPartPath.ResolvePartPath(drawingPart, rel.Target);
            if (partPath == null) return null;

            if (mediaCache.TryGetValue(partPath, out var cached)) return cached;

            var mediaEntry = zip.GetEntry(partPath);
            if (mediaEntry == null) return null;

            var bytes = ReadEntryBytes(mediaEntry);

            mediaCache[partPath] = bytes;
            return bytes;
        }

        /// <summary>读出包条目全部内容（drawing 与 richValue 两条媒体加载路径共用）。</summary>
        private static byte[] ReadEntryBytes(ZipArchiveEntry entry)
        {
            using var src = entry.Open();
            using var ms = new MemoryStream((int)Math.Min(entry.Length, int.MaxValue));
            src.CopyTo(ms);
            return ms.ToArray();
        }

        private static XmlDocument? LoadXml(ZipArchiveEntry entry)
        {
            try
            {
                using var stream = entry.Open();
                var doc = new XmlDocument { XmlResolver = null };
                doc.Load(stream);
                return doc;
            }
            catch (Exception ex)
            {
                // 部件损坏/非法 XML：该部件按缺失处理，图片索引少一部分而非整体失败
                Debug.WriteLine($"解析包部件 {entry.FullName} 失败: {ex.Message}");
                return null;
            }
        }
    }
}
