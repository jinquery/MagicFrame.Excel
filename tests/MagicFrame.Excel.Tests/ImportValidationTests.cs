using MagicFrame.Excel.Engine;
using MagicFrame.Excel.Exceptions;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using NPOI.XSSF.UserModel;
using Xunit;

namespace MagicFrame.Excel.Tests;

/// <summary>
/// 导入时列对应关系校验测试：
/// 缺失列、多余列、只读列忽略、表头重复列
/// </summary>
public class ImportValidationTests
{
    private readonly ExcelEngine _engine = TestHelper.Engine;

    /// <summary>把导出结果按委托改造后再序列化为字节（用于构造异常表）</summary>
    private static byte[] Mutate(ExcelExportResult result, Action<XSSFSheet> mutate)
    {
        var sheet = (XSSFSheet)result.Workbook.GetSheetAt(0);
        mutate(sheet);
        using var ms = new MemoryStream();
        result.Workbook.Write(ms, true);
        return ms.ToArray();
    }

    // ---------- 缺失列 ----------

    [Fact]
    public void Missing_Column_Throws_By_Default()
    {
        using var result = _engine.Export(TestHelper.SampleEmployees(), new ExcelExportOptions { SheetName = "员工表" });

        // 把"年龄"表头清空 -> 该列缺失
        var bytes = Mutate(result, sheet => sheet.GetRow(0).GetCell(2).SetCellValue(""));

        var ex = Assert.Throws<ExcelImportException>(() =>
            _engine.Import<Employee>(bytes, new ExcelImportOptions { SheetName = "员工表" }));

        Assert.Contains("缺少列", ex.Message);
        Assert.Contains("年龄", ex.Message);
    }

    [Fact]
    public void Missing_Column_DoesNot_Throw_When_Disabled_And_Reports_Error()
    {
        using var result = _engine.Export(TestHelper.SampleEmployees(), new ExcelExportOptions { SheetName = "员工表" });

        var bytes = Mutate(result, sheet => sheet.GetRow(0).GetCell(2).SetCellValue(""));

        var options = new ExcelImportOptions { SheetName = "员工表", ThrowOnMissingColumns = false };
        var imported = _engine.Import<Employee>(bytes, options);

        // 缺失列不参与写入，年龄保持默认 0
        Assert.Equal(3, imported.Count);
        Assert.All(imported, e => Assert.Equal(0, e.Age));
        Assert.Contains(options.Errors, e => e.Contains("缺少列") && e.Contains("年龄"));
    }

    // ---------- 多余列 ----------

    [Fact]
    public void Unexpected_Column_Is_Reported_Not_Thrown_By_Default()
    {
        using var result = _engine.Export(TestHelper.SampleEmployees(), new ExcelExportOptions { SheetName = "员工表" });

        // 追加一列"神秘列"
        var bytes = Mutate(result, sheet =>
        {
            sheet.GetRow(0).CreateCell(7).SetCellValue("神秘列");
            sheet.GetRow(1).CreateCell(7).SetCellValue("x");
        });

        var options = new ExcelImportOptions { SheetName = "员工表" };
        var imported = _engine.Import<Employee>(bytes, options);

        Assert.Equal(3, imported.Count);
        Assert.Contains(options.Errors, e => e.Contains("无法对应实体列") && e.Contains("神秘列"));
    }

    [Fact]
    public void Unexpected_Column_Throws_When_Enabled()
    {
        using var result = _engine.Export(TestHelper.SampleEmployees(), new ExcelExportOptions { SheetName = "员工表" });

        var bytes = Mutate(result, sheet =>
        {
            sheet.GetRow(0).CreateCell(7).SetCellValue("神秘列");
            sheet.GetRow(1).CreateCell(7).SetCellValue("x");
        });

        var options = new ExcelImportOptions { SheetName = "员工表", ThrowOnUnexpectedColumns = true };
        var ex = Assert.Throws<ExcelImportException>(() => _engine.Import<Employee>(bytes, options));
        Assert.Contains("无法对应实体列", ex.Message);
        Assert.Contains("神秘列", ex.Message);
    }

    // ---------- 只读列 ----------

    [Fact]
    public void ReadOnly_Column_Is_Ignored_By_Default()
    {
        var rows = new[] { new ReadOnlyEntity { Name = "张三" } };
        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "只读表" });

        // 清空只读列表头 -> 缺列但应被忽略，不抛异常
        var bytes = Mutate(result, sheet => sheet.GetRow(0).GetCell(1).SetCellValue(""));

        var options = new ExcelImportOptions { SheetName = "只读表" };
        var imported = _engine.Import<ReadOnlyEntity>(bytes, options);

        Assert.Single(imported);
        Assert.Equal("张三", imported[0].Name);
        Assert.Contains(options.Errors, e => e.Contains("已忽略") && e.Contains("只读计算值"));
    }

    [Fact]
    public void ReadOnly_Column_Is_Required_When_Not_Ignored()
    {
        var rows = new[] { new ReadOnlyEntity { Name = "张三" } };
        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "只读表" });

        var bytes = Mutate(result, sheet => sheet.GetRow(0).GetCell(1).SetCellValue(""));

        var options = new ExcelImportOptions { SheetName = "只读表", IgnoreReadOnlyColumns = false };
        var ex = Assert.Throws<ExcelImportException>(() => _engine.Import<ReadOnlyEntity>(bytes, options));
        Assert.Contains("缺少列", ex.Message);
    }

    [Fact]
    public void ReadOnly_Column_Is_Ignored_On_Import_Without_Error()
    {
        var rows = new[] { new ReadOnlyEntity { Name = "张三" } };
        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "只读表" });

        var options = new ExcelImportOptions { SheetName = "只读表" };
        var imported = _engine.Import<ReadOnlyEntity>(bytesOf(result), options);

        Assert.Single(imported);
        Assert.Equal("张三", imported[0].Name);
        Assert.Equal(42, imported[0].ReadOnlyValue); // get-only 始终返回计算值
    }

    // ---------- 表头重复列 ----------

    [Fact]
    public void Duplicate_Header_Is_Reported_And_Maps_First()
    {
        using var result = _engine.Export(TestHelper.SampleEmployees(), new ExcelExportOptions { SheetName = "员工表" });

        // 在末尾追加一列，列名与"姓名"(A列) 重复 -> 表头重复，无列缺失
        var bytes = Mutate(result, sheet =>
        {
            sheet.GetRow(0).CreateCell(7).SetCellValue("姓名");
            sheet.GetRow(1).CreateCell(7).SetCellValue("xx");
        });

        var options = new ExcelImportOptions { SheetName = "员工表" };
        var imported = _engine.Import<Employee>(bytes, options);

        Assert.Equal(3, imported.Count);
        Assert.Equal("张三", imported[0].Name);          // 姓名映射到首个"姓名"列（A 列）
        Assert.Equal("研发", imported[0].Department);    // 其余列不受影响
        Assert.Contains(options.Errors, e => e.Contains("重复列名"));
    }

    // ---------- 属性不存在 ----------

    [Fact]
    public void Column_With_Unknown_Field_Imports_Without_Error()
    {
        // 程序化列定义中 Field 指向不存在的属性：
        // 视为可写（自定义访问器可能处理任意字段），写入时反射访问器静默跳过，不抛异常
        var columns = new[]
        {
            new ExcelColumn { Name = "编号", Field = "Id", Order = 1 },
            new ExcelColumn { Name = "幽灵", Field = "NoSuchProperty", Order = 2 },
        };
        var rows = new[] { new ExtensionTests.SimpleRow { Id = 1, Name = "A" } };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表", Columns = columns });

        var options = new ExcelImportOptions { SheetName = "表", Columns = columns };
        var imported = _engine.Import<ExtensionTests.SimpleRow>(bytesOf(result), options);

        Assert.Single(imported);
        Assert.Equal(1, imported[0].Id);
        Assert.DoesNotContain(options.Errors, e => e.Contains("缺少列"));
    }

    private static byte[] bytesOf(ExcelExportResult result)
    {
        using var ms = new MemoryStream();
        result.Workbook.Write(ms, true);
        return ms.ToArray();
    }
}
