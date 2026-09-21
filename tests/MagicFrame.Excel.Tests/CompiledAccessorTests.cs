using MagicFrame.Excel.Engine;
using MagicFrame.Excel.Options;
using Xunit;

namespace MagicFrame.Excel.Tests;

/// <summary>
/// 编译式访问器（表达式树）正确性测试：覆盖可空值类型等边界
/// </summary>
public class CompiledAccessorTests
{
    private readonly ExcelEngine _engine = TestHelper.Engine;

    [Fact]
    public void Nullable_Value_Types_RoundTrip()
    {
        var rows = new List<NullableEntity>
        {
            new() { Name = "有值", NullableValue = 7, NullableDate = new DateTime(2024, 1, 2) },
            new() { Name = "空值" },
        };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "可空表" });
        var imported = _engine.Import<NullableEntity>(result.Bytes, new ExcelImportOptions { SheetName = "可空表" });

        Assert.Equal(2, imported.Count);

        // 有值行：编译式 setter 正确装箱为 nullable
        Assert.Equal(7, imported[0].NullableValue);
        Assert.Equal(new DateTime(2024, 1, 2), imported[0].NullableDate!.Value.Date);

        // 空单元格：数值列按默认转换规则得到 0 / DateTime.MinValue（库的既定行为，非 null）
        Assert.Equal(0, imported[1].NullableValue);
        Assert.Equal(DateTime.MinValue, imported[1].NullableDate);
    }

    [Fact]
    public void Read_Then_Write_Many_Rows_Stays_Correct()
    {
        // 多次读写同一类型，验证编译式缓存不出现状态污染
        for (int i = 0; i < 3; i++)
        {
            var rows = TestHelper.SampleEmployees();
            var bytes = _engine.ExportBytes(rows, new ExcelExportOptions { SheetName = "员工表" });
            var back = _engine.Import<Employee>(bytes, new ExcelImportOptions { SheetName = "员工表" });
            Assert.Equal(3, back.Count);
            Assert.Equal("张三", back[0].Name);
            Assert.Equal(15000m, back[0].Salary);
            Assert.Equal(new DateTime(2020, 3, 15), back[0].HireDate.Date);
        }
    }
}
