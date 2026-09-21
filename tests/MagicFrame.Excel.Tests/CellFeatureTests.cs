using MagicFrame.Excel.Engine;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using Xunit;

namespace MagicFrame.Excel.Tests;

/// <summary>
/// 数据验证（下拉）、0 值空白、隐藏列测试
/// </summary>
public class CellFeatureTests
{
    private readonly ExcelEngine _engine = TestHelper.Engine;

    // ---------- 下拉验证 ----------

    [Fact]
    public void Dropdown_Validation_Is_Applied_To_Sheet()
    {
        var rows = TestHelper.SampleEmployees();
        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "员工表" });

        var sheet = result.Workbook.GetSheetAt(0);
        var validations = sheet.GetDataValidations();

        // 部门列带下拉（B 列），至少 1 条验证
        Assert.NotEmpty(validations);

        var deptValidation = validations.FirstOrDefault(v => RegionsContainColumn(v.Regions, 1));
        Assert.NotNull(deptValidation);
    }

    [Fact]
    public void Validation_Applies_To_New_Rows_Within_MaxRow()
    {
        var rows = TestHelper.SampleEmployees();
        var options = new ExcelExportOptions { SheetName = "员工表", ValidationMaxRow = 50 };

        using var result = _engine.Export(rows, options);
        var sheet = result.Workbook.GetSheetAt(0);

        var validation = sheet.GetDataValidations().First(v => RegionsContainColumn(v.Regions, 1));
        var region = validation.Regions.GetCellRangeAddress(0);
        Assert.Equal(1, region.FirstRow);          // 从首行数据开始
        Assert.True(region.LastRow >= 49);          // 覆盖到第 50 行
    }

    // ---------- 0 值空白 ----------

    [Fact]
    public void Zero_Value_Is_Written_As_Empty_Cell()
    {
        var rows = new List<ZeroBlankItem>
        {
            new() { Name = "A", Value = 0 },
            new() { Name = "B", Value = 3.14 },
        };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表" });
        var sheet = result.Workbook.GetSheetAt(0);

        // 0 值写为空白字符串
        var zeroCell = sheet.GetRow(1).GetCell(1);
        Assert.Equal(CellType.String, zeroCell.CellType);
        Assert.Equal("", zeroCell.StringCellValue);

        // 非 0 值按数值写入
        var valueCell = sheet.GetRow(2).GetCell(1);
        Assert.Equal(CellType.Numeric, valueCell.CellType);
        Assert.Equal(3.14, valueCell.NumericCellValue, 4);
    }

    // ---------- 隐藏列 ----------

    [Fact]
    public void Invisible_Column_Is_Hidden_But_Present()
    {
        var rows = new List<HiddenColumnItem>
        {
            new() { VisibleCol = "A", HiddenCol = "secret" },
        };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表" });
        var sheet = result.Workbook.GetSheetAt(0);

        Assert.True(sheet.IsColumnHidden(1)); // B 列隐藏
        Assert.False(sheet.IsColumnHidden(0)); // A 列可见
    }

    // ---------- 冻结窗格 ----------

    [Fact]
    public void FreezeHeader_Freezes_Header_Rows()
    {
        using var result = _engine.Export(TestHelper.SampleEmployees(), new ExcelExportOptions { SheetName = "员工表" });
        var sheet = Assert.IsType<NPOI.XSSF.UserModel.XSSFSheet>(result.Workbook.GetSheetAt(0));

        var pane = sheet.PaneInformation;
        Assert.NotNull(pane);
        Assert.True(pane.IsFreezePane());
        Assert.Equal(1, pane.HorizontalSplitTopRow); // 单表头冻结 1 行
    }

    // ---------- 工具 ----------

    private static bool RegionsContainColumn(CellRangeAddressList regions, int column)
    {
        for (int i = 0; i < regions.CountRanges(); i++)
        {
            if (regions.GetCellRangeAddress(i).FirstColumn == column)
            {
                return true;
            }
        }
        return false;
    }
}
