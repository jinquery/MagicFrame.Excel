using MagicFrame.Excel.Options;

namespace MagicFrame.Excel.Engine;

/// <summary>
/// 多 Sheet 导出规格：描述一个 Sheet 的数据与类型
/// </summary>
public class SheetSpec
{
    /// <summary>Sheet 名</summary>
    public string Name { get; set; } = "";

    /// <summary>行数据类型（用于解析列定义）</summary>
    public Type RowType { get; set; } = null!;

    /// <summary>行数据集合（object 装箱）</summary>
    public IEnumerable<object> Rows { get; set; } = Array.Empty<object>();

    /// <summary>该 Sheet 的导出选项；为空时使用导出方法的默认选项</summary>
    public ExcelExportOptions? Options { get; set; }

    public static SheetSpec Of<T>(string name, IEnumerable<T> rows, ExcelExportOptions? options = null)
        where T : new()
    {
        return new SheetSpec
        {
            Name = name,
            RowType = typeof(T),
            Rows = rows.Cast<object>(),
            Options = options
        };
    }
}
