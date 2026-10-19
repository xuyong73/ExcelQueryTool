using System.Collections;
using System.Collections.Specialized;

namespace ExcelQueryTool.Services
{
    /// <summary>
    /// 轻量行集合：替代 DataTable/DataView 作为 DataGrid 数据源（2026-09-28 路线 A）。
    ///
    /// 行数据沿用解析管线产出的 <c>object[]</c>（下标 0 = RowNumber 序号，单元格均为
    /// 显示字符串或 <see cref="ImageTextPair"/>），本类只补三件事：
    /// ① 列名 → 下标映射，供 <see cref="RowView"/> 的字符串索引器绑定（表头可含 <c>/</c> 等路径特殊字符）；
    /// ② <b>惰性</b>行视图缓存 —— Standard 虚拟化下容器不复用，每次滚动进入视口都要取行项，
    ///    缓存保证每行只物化一次 RowView（DataView 则在绑定瞬间物化全表 DataRowView = rebind 1.5s 的来源）；
    /// ③ 预览阶段的逐行变更通知（<c>raiseChanges: true</c> 时），支撑「前 N 行边加载边显示」；
    ///    批量段所在的全量表<b>不传通知</b> —— 全程通知成本实测 6s→10s（见 MainWindow 双表注释）。
    ///
    /// 性能契约：Add/Clear 均摊 O(1)（只动 List + 视图数组按需倍增）；
    /// this[int] 首次访问物化 RowView 并缓存（数组 8B 引用/行；最坏全表滚一遍多出 ~24B/行，
    /// 仍远小于 DataRow+DataRowView 双份固定开销）。线程约束：Add/Clear 与视图物化都在 UI 线程。
    /// </summary>
    public sealed class RowList : IList, INotifyCollectionChanged
    {
        private List<object[]> _rows = [];
        private RowView?[] _views = [];
        private Dictionary<string, int> _map = new(StringComparer.Ordinal);
        private readonly bool _raiseChanges;

        /// <param name="raiseChanges">绑给界面、需要边灌边显示的表传 true（预览表）；全量表传 false</param>
        public RowList(bool raiseChanges = false) => _raiseChanges = raiseChanges;

        public event NotifyCollectionChangedEventHandler? CollectionChanged;

        /// <summary>
        /// 重建列映射并清空数据。RowNumber 固定占下标 0（与解析管线的行布局一致），
        /// <paramref name="headers"/> 是数据列（下标 1 起）；重名列由上游 ColumnNaming 去重（重名→_2）。
        /// </summary>
        public void Configure(IEnumerable<string> headers)
        {
            _rows = [];
            _views = [];
            _map = new Dictionary<string, int>(StringComparer.Ordinal) { ["RowNumber"] = 0 };
            int index = 1;
            foreach (var header in headers)
                _map[header] = index++;
        }

        public void Add(object[] row)
        {
            _rows.Add(row);
            if (_views.Length < _rows.Count)
                Array.Resize(ref _views, Math.Max(64, _views.Length * 2));

            if (_raiseChanges)
                CollectionChanged?.Invoke(this, new NotifyCollectionChangedEventArgs(
                    NotifyCollectionChangedAction.Add, ViewAt(_rows.Count - 1), _rows.Count - 1));
        }

        public void Clear()
        {
            bool hadRows = _rows.Count > 0;
            _rows.Clear();
            Array.Clear(_views);
            if (_raiseChanges && hadRows)
                CollectionChanged?.Invoke(this, new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
        }

        private RowView ViewAt(int index)
        {
            var view = _views[index];
            if (view == null)
            {
                view = new RowView(_map, _rows[index], index, this);
                _views[index] = view;
            }
            return view;
        }

        // ── IList：ListCollectionView 需要的只读面 ──
        // IsReadOnly/IsFixedSize=true 让 CollectionView 走纯展示路径（无 AddNew/编辑）；
        // 灌数用本类自己的 Add/Clear，不走接口成员。

        public int Count => _rows.Count;
        public bool IsReadOnly => true;
        public bool IsFixedSize => true;
        public bool IsSynchronized => false;
        public object SyncRoot => ((ICollection)_rows).SyncRoot;

        // IList 空性注解 get（非空）/set（可空）两侧不一致，单一索引器签名无法同时精确匹配，
        // 隐式实现取 object? 并就地抑制 —— 刻意决定，勿当疏漏清理
#pragma warning disable CS8767, CS8769
        public object? this[int index]
        {
            get => ViewAt(index);
            set => throw new NotSupportedException();
        }
#pragma warning restore CS8767, CS8769

        public int IndexOf(object? item) => item is RowView v && ReferenceEquals(v.Owner, this) ? v.Index : -1;
        public bool Contains(object? item) => IndexOf(item) >= 0;

        public IEnumerator GetEnumerator()
        {
            // 惰性迭代：ListCollectionView 的局部遍历（选中项定位等）只触碰访问到的行
            for (int i = 0; i < _rows.Count; i++)
                yield return ViewAt(i);
        }

        void IList.Insert(int index, object? value) => throw new NotSupportedException();
        void IList.Remove(object? value) => throw new NotSupportedException();
        void IList.RemoveAt(int index) => throw new NotSupportedException();
        int IList.Add(object? value) => throw new NotSupportedException();
        public void CopyTo(Array array, int arrayIndex)
        {
            for (int i = 0; i < _rows.Count; i++)
                array.SetValue(ViewAt(i), arrayIndex + i);
        }
    }

    /// <summary>
    /// 单行的绑定视图：字符串索引器供 <c>[header]</c> 绑定路径取值（DataView 字符串索引器的对等物），
    /// <see cref="RowNumber"/> 是 CLR 属性（行头 <c>{Binding RowNumber}</c> 走属性名而非索引器）、
    /// <see cref="Values"/> 供行高统计遍历整行。实例由 <see cref="RowList"/> 惰性物化并缓存。
    /// </summary>
    public sealed class RowView
    {
        private readonly Dictionary<string, int> _map;
        private readonly object?[] _values;

        internal RowView(Dictionary<string, int> map, object?[] values, int index, RowList owner)
        {
            _map = map;
            _values = values;
            Index = index;
            Owner = owner;
        }

        internal RowList Owner { get; }
        internal int Index { get; }

        /// <summary>下标 0 的行序号（第 1 条数据 = 1）；缺失或非 int 返回 0</summary>
        public int RowNumber => _values is { Length: > 0 } && _values[0] is int no ? no : 0;

        /// <summary>原始行数组（含下标 0 的 RowNumber）</summary>
        public object?[] Values => _values;

        /// <summary>按列名取单元格；未知列名或越界返回 null（对齐 DBNull 绑定显示为空的行为）</summary>
        public object? this[string header]
            => _map.TryGetValue(header, out int i) && i < _values.Length ? _values[i] : null;
    }
}
