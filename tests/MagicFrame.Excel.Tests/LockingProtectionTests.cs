using MagicFrame.Excel.Engine;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace MagicFrame.Excel.Tests;

/// <summary>
/// 整列锁定 + 工作表保护测试
/// </summary>
public class LockingProtectionTests
{
    private readonly ExcelEngine _engine = TestHelper.Engine;

    [Fact]
    public void Locked_Columns_Are_Locked_Unlocked_Are_Not()
    {
        var rows = new List<Account>
        {
            new() { Login = "admin", Password = "123456", Name = "管理员" },
        };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "账号表" });
        var sheet = result.Workbook.GetSheetAt(0);
        var dataRow = sheet.GetRow(1);

        // 账号列（IsLocked=true）锁定
        Assert.True(dataRow.GetCell(0).CellStyle.IsLocked);
        // 密码列（IsLocked=true）锁定
        Assert.True(dataRow.GetCell(1).CellStyle.IsLocked);
        // 姓名列（未锁定）可编辑
        Assert.False(dataRow.GetCell(2).CellStyle.IsLocked);
    }

    [Fact]
    public void Sheet_Is_Protected_With_Defaults_When_Locked_Column_Exists()
    {
        var rows = new List<Account>
        {
            new() { Login = "admin", Password = "123456", Name = "管理员" },
        };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "账号表" });
        var sheet = result.Workbook.GetSheetAt(0);

        Assert.True(sheet.Protect);

        var xssfSheet = Assert.IsType<XSSFSheet>(sheet);
        var protection = xssfSheet.GetCTWorksheet()?.sheetProtection;
        Assert.NotNull(protection);

        // 默认放开 新增行 / 删除行 / 调整格式 / 排序 / 自动筛选
        Assert.False(protection.insertRows);
        Assert.False(protection.deleteRows);
        Assert.False(protection.formatCells);
        Assert.False(protection.formatColumns);
        Assert.False(protection.formatRows);
        Assert.False(protection.sort);
        Assert.False(protection.autoFilter);
    }

    [Fact]
    public void Protection_Can_Be_Disabled_Via_Options()
    {
        var rows = new List<Account>
        {
            new() { Login = "admin", Password = "123456", Name = "管理员" },
        };

        var options = new ExcelExportOptions
        {
            SheetName = "账号表",
            Protection = new SheetProtectionOptions { Enabled = false },
        };

        using var result = _engine.Export(rows, options);
        Assert.False(result.Workbook.GetSheetAt(0).Protect);
    }

    [Fact]
    public void Protection_Allows_Controlling_Insert_And_Delete()
    {
        var rows = new List<Account>
        {
            new() { Login = "admin", Password = "123456", Name = "管理员" },
        };

        var options = new ExcelExportOptions
        {
            SheetName = "账号表",
            Protection = new SheetProtectionOptions { AllowInsertRows = false, AllowDeleteRows = false },
        };

        using var result = _engine.Export(rows, options);
        var xssfSheet = Assert.IsType<XSSFSheet>(result.Workbook.GetSheetAt(0));
        var protection = xssfSheet.GetCTWorksheet()!.sheetProtection;

        // 关闭时 insertRows 保持 true（即禁止插入）
        Assert.True(protection.insertRows);
        Assert.True(protection.deleteRows);
        // 未关闭的 sort / autoFilter 仍放开
        Assert.False(protection.sort);
        Assert.False(protection.autoFilter);
    }

    [Fact]
    public void Sort_And_AutoFilter_Can_Be_Blocked_Via_Options()
    {
        var rows = new List<Account>
        {
            new() { Login = "admin", Password = "123456", Name = "管理员" },
        };

        var options = new ExcelExportOptions
        {
            SheetName = "账号表",
            Protection = new SheetProtectionOptions { AllowSort = false, AllowAutoFilter = false },
        };

        using var result = _engine.Export(rows, options);
        var xssfSheet = Assert.IsType<XSSFSheet>(result.Workbook.GetSheetAt(0));
        var protection = xssfSheet.GetCTWorksheet()!.sheetProtection;

        // 未放开 -> 保持禁止（true）
        Assert.True(protection.sort);
        Assert.True(protection.autoFilter);
    }

    [Fact]
    public void Exported_File_Allows_Sort_And_AutoFilter_By_Default()
    {
        // 端到端验证：锁定列默认导出的文件，sheetProtection 中 sort/autoFilter 显式为允许(0)。
        var rows = new List<Account>
        {
            new() { Login = "admin", Password = "123456", Name = "管理员" },
        };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "账号表" });
        var bytes = result.Bytes;

        using var workbook = new XSSFWorkbook(new MemoryStream(bytes));
        var protection = Assert.IsType<XSSFSheet>(workbook.GetSheetAt(0)).GetCTWorksheet()!.sheetProtection;
        Assert.NotNull(protection);
        Assert.False(protection.sort);
        Assert.False(protection.autoFilter);
    }
}
