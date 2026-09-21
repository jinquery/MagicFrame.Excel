using MagicFrame.Excel.Attributes;
using NPOI.HSSF.Util;

namespace MagicFrame.Excel.Tests;

/// <summary>基础实体：覆盖常用类型</summary>
public class Employee
{
    [ExcelColumn("姓名", Order = 1, Width = 10)]
    public string Name { get; set; } = "";

    [ExcelColumn("部门", Order = 2, Width = 14, DropdownOptions = new[] { "研发", "测试", "产品", "运营" })]
    public string Department { get; set; } = "";

    [ExcelColumn("年龄", Order = 3, Width = 8)]
    public int Age { get; set; }

    [ExcelColumn("工资", Order = 4, Width = 12)]
    public decimal Salary { get; set; }

    [ExcelColumn("在职", Order = 5, Width = 6)]
    public bool Active { get; set; }

    [ExcelColumn("入职日期", Order = 6, Width = 14)]
    public DateTime HireDate { get; set; }

    [ExcelColumn("备注", Order = 7, Width = 18)]
    public string? Remark { get; set; }
}

/// <summary>公式实体：合计列由 数量*单价 计算</summary>
public class Product
{
    [ExcelColumn("名称", Order = 1)]
    public string Name { get; set; } = "";

    [ExcelColumn("数量", Order = 2)]
    public int Quantity { get; set; }

    [ExcelColumn("单价", Order = 3)]
    public decimal UnitPrice { get; set; }

    [ExcelColumn("合计", Order = 4, Formula = "{Quantity}*{UnitPrice}")]
    public decimal Total { get; set; }
}

/// <summary>锁定实体：密码列锁定，名称列可编辑</summary>
public class Account
{
    [ExcelColumn("账号", Order = 1, IsLocked = true)]
    public string Login { get; set; } = "";

    [ExcelColumn("密码", Order = 2, IsLocked = true)]
    public string Password { get; set; } = "";

    [ExcelColumn("姓名", Order = 3)]
    public string Name { get; set; } = "";
}

/// <summary>分组表头实体：基础信息 + 考核信息 两组</summary>
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

    [ExcelColumn("总分", Order = 5, GroupName = "Kpi", GroupText = "考核信息", Formula = "{Performance}+{Attendance}")]
    public int Total { get; set; }

    [ExcelColumn("评级", Order = 6, DropdownOptions = new[] { "A", "B", "C" })]
    public string Grade { get; set; } = "";
}

/// <summary>0 值空白实体</summary>
public class ZeroBlankItem
{
    [ExcelColumn("名称", Order = 1)]
    public string Name { get; set; } = "";

    [ExcelColumn("数值", Order = 2, ZeroShowWhiteSpace = true)]
    public double Value { get; set; }
}

/// <summary>隐藏列实体</summary>
public class HiddenColumnItem
{
    [ExcelColumn("可见", Order = 1)]
    public string VisibleCol { get; set; } = "";

    [ExcelColumn("隐藏", Order = 2, Visible = false)]
    public string HiddenCol { get; set; } = "";
}

/// <summary>全字符串实体：用于"跳过空行"测试（空行的数值列默认为 0，会被视为有值）</summary>
public class StringRow
{
    [ExcelColumn("名称", Order = 1)]
    public string Name { get; set; } = "";

    [ExcelColumn("城市", Order = 2)]
    public string City { get; set; } = "";
}

/// <summary>只读列实体：含 get-only 属性，用于导入列对应校验测试。
/// 说明：private set 属性经反射可写，不算只读；仅 get-only（CanWrite=false）才视为只读忽略。</summary>
public class ReadOnlyEntity
{
    [ExcelColumn("名称", Order = 1)]
    public string Name { get; set; } = "";

    [ExcelColumn("只读计算值", Order = 2)]
    public int ReadOnlyValue => 42;
}

/// <summary>可空值类型实体：验证编译式 setter 对 nullable 的赋值</summary>
public class NullableEntity
{
    [ExcelColumn("名称", Order = 1)]
    public string Name { get; set; } = "";

    [ExcelColumn("可空数值", Order = 2)]
    public int? NullableValue { get; set; }

    [ExcelColumn("可空日期", Order = 3)]
    public DateTime? NullableDate { get; set; }
}

/// <summary>导入新特性测试实体：源行号回填 / 逐行错误收集（枚举非法值会触发转换异常）</summary>
public class ImportRow
{
    [ExcelColumn("名称", Order = 1)]
    public string Name { get; set; } = "";

    [ExcelColumn("状态", Order = 2)]
    public TestStatus Status { get; set; }

    /// <summary>非列属性：用于 SourceRowProperty 回填</summary>
    public int SourceRow { get; set; }
}

public enum TestStatus
{
    Pending = 0,
    Active = 1,
    Done = 2,
}

/// <summary>程序化列测试实体：数字格式 / 数据验证</summary>
public class ValidationRow
{
    public int Qty { get; set; }
    public double Ratio { get; set; }
    public DateTime Day { get; set; }
    public string Note { get; set; } = "";
}

/// <summary>用于自定义表头（FixedTitle）扩展测试的实体</summary>
public class TitleItem
{
    [ExcelColumn("标题列1", Order = 1)]
    public string Col1 { get; set; } = "";

    [ExcelColumn("标题列2", Order = 2)]
    public string Col2 { get; set; } = "";
}
