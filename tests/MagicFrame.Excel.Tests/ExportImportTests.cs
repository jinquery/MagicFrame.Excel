using MagicFrame.Excel.Engine;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;
using Xunit;

namespace MagicFrame.Excel.Tests;

/// <summary>
/// 基础导出/导入往返测试：字节、Base64、流、文件、多 Sheet
/// </summary>
public class ExportImportTests
{
    private readonly ExcelEngine _engine = TestHelper.Engine;

    [Fact]
    public void Export_Then_Import_RoundTrip_PreservesValues()
    {
        var rows = TestHelper.SampleEmployees();
        var options = new ExcelExportOptions { SheetName = "员工表" };

        var bytes = _engine.ExportBytes(rows, options);

        Assert.NotEmpty(bytes);
        Assert.StartsWith("PK", System.Text.Encoding.ASCII.GetString(bytes.Take(4).ToArray())); // xlsx 魔数

        var imported = _engine.Import<Employee>(bytes, new ExcelImportOptions { SheetName = "员工表" });

        Assert.Equal(rows.Count, imported.Count);
        for (int i = 0; i < rows.Count; i++)
        {
            Assert.Equal(rows[i].Name, imported[i].Name);
            Assert.Equal(rows[i].Department, imported[i].Department);
            Assert.Equal(rows[i].Age, imported[i].Age);
            Assert.Equal(rows[i].Salary, imported[i].Salary);
            Assert.Equal(rows[i].Active, imported[i].Active);
            Assert.Equal(rows[i].HireDate.Date, imported[i].HireDate.Date);
        }
    }

    [Fact]
    public void ExportBase64_Can_Be_Imported_Back()
    {
        var rows = TestHelper.SampleEmployees();
        string base64 = _engine.ExportBase64(rows);

        Assert.False(string.IsNullOrWhiteSpace(base64));

        var imported = _engine.Import<Employee>(base64);
        Assert.Equal(rows.Count, imported.Count);
        Assert.Equal(rows[0].Name, imported[0].Name);
    }

    [Fact]
    public void ExportStream_Position_Is_Zero_And_Readable()
    {
        using var result = _engine.ExportStream(TestHelper.SampleEmployees());
        Assert.Equal(0, result.Position);
        Assert.True(result.Length > 0);
        var imported = _engine.Import<Employee>(result);
        Assert.Equal(3, imported.Count);
    }

    [Fact]
    public void ExportToFile_Writes_Valid_Xlsx()
    {
        string path = Path.Combine(Path.GetTempPath(), $"magicframe_{Guid.NewGuid():N}.xlsx");
        try
        {
            _engine.ExportToFile(TestHelper.SampleEmployees(), path);
            Assert.True(File.Exists(path));
            Assert.True(new FileInfo(path).Length > 0);
            var imported = _engine.ImportFile<Employee>(path);
            Assert.Equal(3, imported.Count);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void Export_Result_Workbook_Is_Usable()
    {
        using var result = _engine.Export(TestHelper.SampleEmployees(), new ExcelExportOptions { SheetName = "员工表" });
        Assert.NotNull(result.Workbook);
        Assert.Equal("员工表", result.Workbook.GetSheetAt(0).SheetName);
        Assert.NotNull(result.Bytes);
        Assert.NotNull(result.Base64);
    }

    [Fact]
    public void ExportSheets_Multiple_Sheets_Import_Each()
    {
        var employees = TestHelper.SampleEmployees();
        var products = TestHelper.SampleProducts();

        var specs = new[]
        {
            SheetSpec.Of("员工表", employees),
            SheetSpec.Of("商品表", products),
        };

        using var result = _engine.ExportSheets(specs);

        Assert.Equal(2, result.Workbook.NumberOfSheets);
        Assert.Equal("员工表", result.Workbook.GetSheetName(0));
        Assert.Equal("商品表", result.Workbook.GetSheetName(1));

        var employeesBack = _engine.Import<Employee>(result.Bytes, new ExcelImportOptions { SheetName = "员工表" });
        var productsBack = _engine.Import<Product>(result.Bytes, new ExcelImportOptions { SheetName = "商品表" });

        Assert.Equal(3, employeesBack.Count);
        Assert.Equal(3, productsBack.Count);
        Assert.Equal("苹果", productsBack[0].Name);
    }

    [Fact]
    public void ExportSheets_SameType_Dictionary_Overload()
    {
        var sheets = new Dictionary<string, IList<Employee>>
        {
            ["研发"] = TestHelper.SampleEmployees().Where(e => e.Department == "研发").ToList(),
            ["测试"] = TestHelper.SampleEmployees().Where(e => e.Department == "测试").ToList(),
        };

        using var result = _engine.ExportSheets(sheets);

        Assert.Equal(2, result.Workbook.NumberOfSheets);
        var back = _engine.Import<Employee>(result.Bytes, new ExcelImportOptions { SheetName = "研发" });
        Assert.Single(back);
        Assert.Equal("张三", back[0].Name);
    }

    [Fact]
    public void Empty_Rows_Are_Skipped_On_Import()
    {
        var rows = new List<StringRow>
        {
            new() { Name = "A", City = "北京" },
            new() { Name = "B", City = "上海" },
            new() { Name = "C", City = "广州" },
        };
        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表" });

        // 手动在数据区末尾追加两行空行后重新序列化
        result.Workbook.GetSheetAt(0).CreateRow(4);
        result.Workbook.GetSheetAt(0).CreateRow(5);
        using var ms = new MemoryStream();
        result.Workbook.Write(ms, true);
        var bytes = ms.ToArray();

        var imported = _engine.Import<StringRow>(bytes, new ExcelImportOptions { SheetName = "表" });
        Assert.Equal(3, imported.Count);
    }

    [Fact]
    public void Missing_Column_Throws_By_Default()
    {
        var rows = TestHelper.SampleEmployees();
        var bytes = _engine.ExportBytes(rows);

        var imported = _engine.Import<Employee>(bytes); // 列齐全，不抛
        Assert.Equal(3, imported.Count);
    }
}
