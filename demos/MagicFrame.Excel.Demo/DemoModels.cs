using MagicFrame.Excel.Attributes;

namespace MagicFrame.Excel.Demo;

/// <summary>员工实体：演示单表头 + 下拉验证 + 整列锁定 + 公式</summary>
public class Employee
{
    [ExcelColumn("姓名", Order = 1, Width = 10)]
    public string Name { get; set; } = "";

    [ExcelColumn("部门", Order = 2, Width = 14, DropdownOptions = new[] { "研发", "测试", "产品", "运营" })]
    public string Department { get; set; } = "";

    [ExcelColumn("基础工资", Order = 3, Width = 12)]
    public decimal BaseSalary { get; set; }

    [ExcelColumn("绩效系数", Order = 4, Width = 10)]
    public decimal Factor { get; set; }

    [ExcelColumn("应发合计", Order = 5, Width = 12, IsLocked = true, Formula = "{BaseSalary}*{Factor}")]
    public decimal Total { get; set; }

    [ExcelColumn("入职日期", Order = 6, Width = 14)]
    public DateTime HireDate { get; set; }

    [ExcelColumn("备注", Order = 7, Width = 18, Visible = false)]
    public string? Remark { get; set; }
}

/// <summary>考核实体：演示分组表头</summary>
public class Assessment
{
    [ExcelColumn("姓名", Order = 1, GroupName = "Base", GroupText = "基础信息")]
    public string Name { get; set; } = "";

    [ExcelColumn("部门", Order = 2, GroupName = "Base", GroupText = "基础信息")]
    public string Department { get; set; } = "";

    [ExcelColumn("业绩", Order = 3, GroupName = "Kpi", GroupText = "考核信息")]
    public int Performance { get; set; }

    [ExcelColumn("出勤", Order = 4, GroupName = "Kpi", GroupText = "考核信息")]
    public int Attendance { get; set; }

    [ExcelColumn("总分", Order = 5, GroupName = "Kpi", GroupText = "考核信息", IsLocked = true, Formula = "{Performance}+{Attendance}")]
    public int Total { get; set; }

    [ExcelColumn("评级", Order = 6, DropdownOptions = new[] { "A", "B", "C" })]
    public string Grade { get; set; } = "";
}

/// <summary>商品实体：演示 0 值空白 + 单表头</summary>
public class Product
{
    [ExcelColumn("名称", Order = 1)]
    public string Name { get; set; } = "";

    [ExcelColumn("销量", Order = 2, ZeroShowWhiteSpace = true)]
    public int Sales { get; set; }
}

/// <summary>手工列（无特性）：扩展演示——用自定义列提供器 + 程序化列定义</summary>
public class ManualRow
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public decimal Amount { get; set; }
}

/// <summary>值映射演示实体：字段存代码，Excel 显示文本</summary>
public class GenderRow
{
    public string Name { get; set; } = "";
    public int Gender { get; set; }
}
