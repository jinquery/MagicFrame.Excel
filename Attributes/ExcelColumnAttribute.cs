using MagicFrame.Excel.Model;
using NPOI.HSSF.Util;

namespace MagicFrame.Excel.Attributes;

/// <summary>
/// 标注在 POCO 属性上，声明该属性对应 Excel 中的一列。
/// 纯 POCO 方式：实体无需继承任何基类。
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public class ExcelColumnAttribute : Attribute
{
    /// <summary>表头显示文字</summary>
    public string Name { get; }

    /// <summary>实体属性名，为空时取属性名</summary>
    public string? Field { get; set; }

    /// <summary>列顺序</summary>
    public int Order { get; set; }

    /// <summary>是否显示</summary>
    public bool Visible { get; set; } = true;

    /// <summary>整列锁定</summary>
    public bool IsLocked { get; set; }

    /// <summary>公式占位符</summary>
    public string? Formula { get; set; }

    /// <summary>0 值显示为空白</summary>
    public bool ZeroShowWhiteSpace { get; set; }

    /// <summary>表头字体颜色</summary>
    public short HeaderColor { get; set; } = HSSFColor.Black.Index;

    /// <summary>分组名称</summary>
    public string? GroupName { get; set; }

    /// <summary>分组显示文字</summary>
    public string? GroupText { get; set; }

    /// <summary>数据验证下拉选项</summary>
    public string[]? DropdownOptions { get; set; }

    /// <summary>数据验证类型</summary>
    public ValidationKind ValidationKind { get; set; } = ValidationKind.None;

    /// <summary>数据验证参数 1（区间最小值 / 自定义公式 / 公式列表来源）</summary>
    public string ValidationFormula1 { get; set; } = "";

    /// <summary>数据验证参数 2（区间最大值，仅 Integer/Decimal/Date 使用）</summary>
    public string ValidationFormula2 { get; set; } = "";

    /// <summary>列宽（字符数），0 表示按表头文字自适应</summary>
    public int Width { get; set; }

    /// <summary>水平对齐</summary>
    public CellAlignment Alignment { get; set; } = CellAlignment.Center;

    /// <summary>是否按数值类型写入单元格</summary>
    public bool WriteAsNumeric { get; set; } = true;

    /// <summary>数字格式掩码，如 "0.00%" / "#,##0.00"；为空用默认</summary>
    public string NumberFormat { get; set; } = "";

    /// <summary>
    /// 值映射字符串，格式 "实体值:显示文本"，逗号分隔，如 "0:女,1:男,unknown:2"。
    /// 纯数字键按整数解析，其余按字符串；
    /// 特殊键 <c>unknown</c> 指定"异常值"兜底代码（导入未映射文本时使用），
    /// 未提供时默认加上 <c>unknown:-9999999</c> 表示异常值。
    /// </summary>
    public string? ValueMappings { get; set; }

    public ExcelColumnAttribute(string name)
    {
        Name = name;
    }
}
