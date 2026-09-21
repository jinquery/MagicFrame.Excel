using MagicFrame.Excel.Engine;
using NPOI.SS.UserModel;

namespace MagicFrame.Excel.Tests;

/// <summary>
/// 测试公用工具
/// </summary>
public static class TestHelper
{
    public static ExcelEngine Engine => ExcelEngine.CreateDefault();

    public static List<Employee> SampleEmployees() => new()
    {
        new Employee { Name = "张三", Department = "研发", Age = 28, Salary = 15000m, Active = true, HireDate = new DateTime(2020, 3, 15), Remark = "组长" },
        new Employee { Name = "李四", Department = "测试", Age = 32, Salary = 12000m, Active = true, HireDate = new DateTime(2019, 7, 1), Remark = null },
        new Employee { Name = "王五", Department = "产品", Age = 26, Salary = 18000m, Active = false, HireDate = new DateTime(2022, 11, 20), Remark = "实习" },
    };

    public static List<Product> SampleProducts() => new()
    {
        new Product { Name = "苹果", Quantity = 3, UnitPrice = 10.5m },
        new Product { Name = "香蕉", Quantity = 5, UnitPrice = 4.2m },
        new Product { Name = "橙子", Quantity = 2, UnitPrice = 7.8m },
    };

    /// <summary>把工作簿的第一个 Sheet 转成可见文本表（用于断言）</summary>
    public static string[,] DumpSheet(IWorkbook workbook, int sheetIndex = 0)
    {
        var sheet = workbook.GetSheetAt(sheetIndex);
        int rows = sheet.LastRowNum + 1;
        int cols = 0;
        for (int r = 0; r < rows; r++)
        {
            cols = Math.Max(cols, sheet.GetRow(r)?.LastCellNum ?? 0);
        }
        var grid = new string[rows, cols];
        for (int r = 0; r < rows; r++)
        {
            var row = sheet.GetRow(r);
            if (row == null)
            {
                continue;
            }
            for (int c = 0; c < cols; c++)
            {
                var cell = row.GetCell(c);
                grid[r, c] = cell?.ToString() ?? "";
            }
        }
        return grid;
    }
}
