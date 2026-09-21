using NPOI.HSSF.Util;

namespace MagicFrame.Excel.Model;

/// <summary>
/// 单元格水平对齐方式（与 NPOI 解耦，便于在非 NPOI 环境配置）
/// </summary>
public enum CellAlignment
{
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterSelection,
    Distributed
}

/// <summary>
/// 数据验证类型：下拉列表用 <see cref="ExcelColumn.DropdownOptions"/>（等价于 <see cref="List"/> 的简写）
/// </summary>
public enum ValidationKind
{
    /// <summary>不校验</summary>
    None,

    /// <summary>显式列表（等价于 DropdownOptions）</summary>
    List,

    /// <summary>整数区间（ValidationFormula1/2 为最小/最大值）</summary>
    Integer,

    /// <summary>小数区间（ValidationFormula1/2 为最小/最大值）</summary>
    Decimal,

    /// <summary>日期区间（ValidationFormula1/2 为最小/最大日期，如 "2020-01-01"）</summary>
    Date,

    /// <summary>自定义公式（ValidationFormula1 为公式）</summary>
    CustomFormula,

    /// <summary>公式来源列表（ValidationFormula1 为区域或逗号分隔列表）</summary>
    FormulaList,
}

/// <summary>
/// 列定义：是生成与解析 Excel 的"单一事实来源"。
/// 可通过 <see cref="Attributes.ExcelColumnAttribute"/> 自动发现，也可用流式配置手工构建。
/// </summary>
public class ExcelColumn
{
    /// <summary>表头显示文字（中文列名）</summary>
    public string Name { get; set; } = "";

    /// <summary>实体属性名（T 上的属性/字段）</summary>
    public string Field { get; set; } = "";

    /// <summary>列顺序</summary>
    public int Order { get; set; }

    /// <summary>是否显示，false 时列为隐藏列（保留位置）</summary>
    public bool Visible { get; set; } = true;

    /// <summary>整列锁定：true 时该列（含用户新增行的该列）不可编辑</summary>
    public bool IsLocked { get; set; }

    /// <summary>
    /// 单元格公式占位符，如 "IFERROR({HardSoilLen}+{SoftSoilLen},\"\")"，无需等号，
    /// {} 内为列 Field，生成时替换为当前行的单元格地址（如 B3）。
    /// </summary>
    public string Formula { get; set; } = "";

    /// <summary>0 值显示为空白</summary>
    public bool ZeroShowWhiteSpace { get; set; }

    /// <summary>表头字体颜色（HSSFColor 索引）</summary>
    public short HeaderColor { get; set; } = HSSFColor.Black.Index;

    /// <summary>分组表头：分组名称（同一分组的多列在第 1 行合并显示 GroupText）</summary>
    public string GroupName { get; set; } = "";

    /// <summary>分组表头：分组显示文字</summary>
    public string GroupText { get; set; } = "";

    /// <summary>数据验证下拉选项；为空则不生成下拉</summary>
    public IList<string>? DropdownOptions { get; set; }

    /// <summary>数据验证类型；<see cref="ValidationKind.List"/> 与 DropdownOptions 等价</summary>
    public ValidationKind ValidationKind { get; set; } = ValidationKind.None;

    /// <summary>数据验证参数 1（区间最小值 / 自定义公式 / 公式列表来源）</summary>
    public string ValidationFormula1 { get; set; } = "";

    /// <summary>数据验证参数 2（区间最大值，仅 Integer/Decimal/Date 使用）</summary>
    public string ValidationFormula2 { get; set; } = "";

    /// <summary>列宽（字符数）；为空时按表头文字自适应</summary>
    public int? Width { get; set; }

    /// <summary>水平对齐</summary>
    public CellAlignment Alignment { get; set; } = CellAlignment.Center;

    /// <summary>是否按数值类型写入单元格（false 时统一按文本写入）</summary>
    public bool WriteAsNumeric { get; set; } = true;

    /// <summary>数字格式掩码，如 "0.00%" / "#,##0.00"；为空用默认</summary>
    public string NumberFormat { get; set; } = "";

    /// <summary>
    /// 复制一份（避免引用共享导致的意外修改）
    /// </summary>
    public ExcelColumn Clone() => (ExcelColumn)MemberwiseClone();
}

/// <summary>
/// 表头类型常量：注册/选择对应的写入与读取策略
/// </summary>
public static class HeaderKinds
{
    /// <summary>单行表头</summary>
    public const string Single = "Single";

    /// <summary>多级/分组表头（二级）</summary>
    public const string Group = "Group";
}
