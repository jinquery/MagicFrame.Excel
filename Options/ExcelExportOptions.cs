using MagicFrame.Excel.Model;

namespace MagicFrame.Excel.Options;

/// <summary>
/// 导出选项
/// </summary>
public class ExcelExportOptions
{
    /// <summary>Sheet 名</summary>
    public string SheetName { get; set; } = "sheet1";

    /// <summary>表头类型，默认单表头；可选 <see cref="HeaderKinds.Group"/></summary>
    public string HeaderKind { get; set; } = HeaderKinds.Single;

    /// <summary>
    /// 列定义；为空时由列提供器（默认特性发现）从实体类型解析。
    /// 可手工传 <see cref="ExcelColumn"/> 集合实现流式/自定义配置。
    /// </summary>
    public IReadOnlyList<ExcelColumn>? Columns { get; set; }

    /// <summary>
    /// 保护选项；为 null 时若存在 IsLocked 列则按默认保护（放开新增/删除/格式）执行
    /// </summary>
    public SheetProtectionOptions? Protection { get; set; }

    /// <summary>冻结表头（前 HeaderRowCount 行）</summary>
    public bool FreezeHeader { get; set; } = true;

    /// <summary>冻结前 N 列（0 不冻结；与 FreezeHeader 可同时生效）</summary>
    public int FreezeColumns { get; set; }

    /// <summary>是否在表头行添加自动筛选（AutoFilter）</summary>
    public bool AutoFilter { get; set; }

    /// <summary>数据行高（磅，0 为自动）</summary>
    public double RowHeight { get; set; }

    /// <summary>交替行填充色（HSSFColor 索引，0 不启用），奇数数据行使用</summary>
    public short AlternateRowFillColor { get; set; }

    /// <summary>是否给数据单元格加细边框</summary>
    public bool EnableBorders { get; set; }

    /// <summary>导出进度回调（0~1）</summary>
    public IProgress<double>? Progress { get; set; }

    /// <summary>已知行数（引擎在可计数时预填，供进度回调使用）</summary>
    internal int? KnownRowCount { get; set; }

    /// <summary>数据验证（下拉）作用到的最大行号，默认 1000</summary>
    public int ValidationMaxRow { get; set; } = 1000;

    /// <summary>自适应列宽上限（字符数），默认 25</summary>
    public int MaxColumnWidth { get; set; } = 25;

    /// <summary>
    /// 导出时是否由 NPOI 预计算公式结果并写入缓存值，默认 true。
    /// <para>
    /// true：导出后文件直接可见公式结果；但公式较多时 <c>EvaluateAllFormulaCells</c>
    /// 会对整簿所有公式单元格求值，导出较慢。
    /// </para>
    /// <para>
    /// false：跳过 NPOI 求值，改为标记"打开时全量重算"（fullCalcOnLoad），
    /// 由 Excel 打开时自动计算，导出显著提速；代价是未用 Excel 打开前，
    /// 其它程序读取到的公式缓存值为空（本库导入仍会按需求值）。
    /// </para>
    /// </summary>
    public bool CalculateFormulasOnExport { get; set; } = true;
}
