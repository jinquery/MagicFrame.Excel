using MagicFrame.Excel.Engine;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace MagicFrame.Excel.Tests;

/// <summary>
/// 公式支持测试：生成时写入公式并强制计算，导入时按公式求值
/// </summary>
public class FormulaTests
{
    private readonly ExcelEngine _engine = TestHelper.Engine;

    [Fact]
    public void Export_Writes_Formula_And_Evaluates_Value()
    {
        var rows = TestHelper.SampleProducts();
        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "商品表" });

        var sheet = result.Workbook.GetSheetAt(0);
        var totalCell = sheet.GetRow(1).GetCell(3); // 第 1 行数据，合计列（D2）

        Assert.Equal(CellType.Formula, totalCell.CellType);
        Assert.Equal("B2*C2", totalCell.CellFormula);

        // 引擎已调用 EvaluateAllFormulaCells，缓存值应为 3 * 10.5 = 31.5
        Assert.Equal(31.5, totalCell.NumericCellValue, 4);
    }

    [Fact]
    public void Import_Evaluates_Formula_Column()
    {
        var rows = TestHelper.SampleProducts();
        var bytes = _engine.ExportBytes(rows);

        var imported = _engine.Import<Product>(bytes);

        Assert.Equal(3, imported.Count);
        Assert.Equal(3m * 10.5m, imported[0].Total);
        Assert.Equal(5m * 4.2m, imported[1].Total);
        Assert.Equal(2m * 7.8m, imported[2].Total);
    }

    [Fact]
    public void Formula_With_Group_Header_Works()
    {
        var rows = new List<Assessment>
        {
            new() { Name = "张三", Department = "研发", Performance = 90, Attendance = 95, Grade = "A" },
            new() { Name = "李四", Department = "测试", Performance = 70, Attendance = 80, Grade = "B" },
        };

        var options = new ExcelExportOptions { SheetName = "考核表", HeaderKind = HeaderKinds.Group };
        var bytes = _engine.ExportBytes(rows, options);

        var imported = _engine.Import<Assessment>(bytes, new ExcelImportOptions { SheetName = "考核表", HeaderKind = HeaderKinds.Group });

        Assert.Equal(2, imported.Count);
        Assert.Equal(90 + 95, imported[0].Total);
        Assert.Equal(70 + 80, imported[1].Total);
    }

    [Fact]
    public void Default_Mode_Precomputes_And_Does_Not_Set_Recalc_Flag()
    {
        var rows = TestHelper.SampleProducts();
        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "商品表" });

        var xssf = Assert.IsType<XSSFWorkbook>(result.Workbook);
        Assert.NotEqual(true, xssf.GetCTWorkbook().calcPr?.fullCalcOnLoad); // 默认已预计算，无需打开时重算
    }

    [Fact]
    public void Deferred_Mode_Skips_Precompute_And_Import_Still_Evaluates()
    {
        var rows = TestHelper.SampleProducts();
        var options = new ExcelExportOptions { SheetName = "商品表", CalculateFormulasOnExport = false };

        using var result = _engine.Export(rows, options);
        var sheet = result.Workbook.GetSheetAt(0);
        var totalCell = sheet.GetRow(1).GetCell(3);

        // 仍是公式单元格，公式未变
        Assert.Equal(CellType.Formula, totalCell.CellType);
        Assert.Equal("B2*C2", totalCell.CellFormula);

        // 已标记"打开时全量重算"（calcPr.fullCalcOnLoad）
        var xssf = Assert.IsType<XSSFWorkbook>(result.Workbook);
        Assert.Equal(true, xssf.GetCTWorkbook().calcPr?.fullCalcOnLoad);

        // 未预计算时单元格无缓存值（0/空），但本库导入仍会按公式重新求值
        var imported = _engine.Import<Product>(result.Bytes, new ExcelImportOptions { SheetName = "商品表" });
        Assert.Equal(3, imported.Count);
        Assert.Equal(3m * 10.5m, imported[0].Total);
        Assert.Equal(5m * 4.2m, imported[1].Total);
        Assert.Equal(2m * 7.8m, imported[2].Total);
    }
}
