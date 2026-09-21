using MagicFrame.Excel.Engine;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;
using Xunit;

namespace MagicFrame.Excel.Tests;

/// <summary>
/// 分组/多级表头测试：写入合并单元格、往返导入、列顺序调整仍可解析
/// </summary>
public class GroupHeaderTests
{
    private readonly ExcelEngine _engine = TestHelper.Engine;

    private static List<Assessment> SampleAssessments() => new()
    {
        new Assessment { Name = "张三", Department = "研发", Performance = 90, Attendance = 95, Total = 185, Grade = "A" },
        new Assessment { Name = "李四", Department = "测试", Performance = 70, Attendance = 80, Total = 150, Grade = "B" },
    };

    private static ExcelExportOptions GroupOptions() =>
        new() { SheetName = "考核表", HeaderKind = HeaderKinds.Group };

    private static ExcelImportOptions GroupImportOptions() =>
        new() { SheetName = "考核表", HeaderKind = HeaderKinds.Group };

    [Fact]
    public void Export_Produces_Two_Header_Rows_And_Merged_Groups()
    {
        using var result = _engine.Export(SampleAssessments(), GroupOptions());
        var sheet = result.Workbook.GetSheetAt(0);

        // 表头占 2 行，数据从第 3 行（索引 2）开始
        Assert.Equal(5, sheet.GetRow(1).LastCellNum); // 列名行 5 列（无分组列名合并到第 1 行）
        Assert.NotNull(sheet.GetRow(2));

        // 基础信息 合并 A1:B1，考核信息 合并 C1:E1，无分组列 评级 上下合并 F1:F2
        var merges = sheet.MergedRegions.Select(r => r.FormatAsString()).ToList();
        Assert.Contains("A1:B1", merges);
        Assert.Contains("C1:E1", merges);
        Assert.Contains("F1:F2", merges);

        // 分组行文字
        Assert.Equal("基础信息", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("考核信息", sheet.GetRow(0).GetCell(2).StringCellValue);
    }

    [Fact]
    public void GroupHeader_RoundTrip_Import()
    {
        var bytes = _engine.ExportBytes(SampleAssessments(), GroupOptions());
        var imported = _engine.Import<Assessment>(bytes, GroupImportOptions());

        Assert.Equal(2, imported.Count);
        Assert.Equal("张三", imported[0].Name);
        Assert.Equal("研发", imported[0].Department);
        Assert.Equal(90, imported[0].Performance);
        Assert.Equal(95, imported[0].Attendance);
        Assert.Equal("A", imported[0].Grade);
    }

    [Fact]
    public void GroupHeader_Import_Works_After_Column_Rearrangement()
    {
        using var result = _engine.Export(SampleAssessments(), GroupOptions());
        var sheet = result.Workbook.GetSheetAt(0);

        // 交换"业绩"(C列) 与 "出勤"(D列)：列名行 + 所有数据行 对调，验证按列名解析不受顺序影响
        SwapCells(sheet, 1, 2, 3); // 列名行
        SwapCells(sheet, 2, 2, 3); // 数据行（张三）
        SwapCells(sheet, 3, 2, 3); // 数据行（李四）

        using var ms = new MemoryStream();
        result.Workbook.Write(ms, true);
        var bytes = ms.ToArray();

        var imported = _engine.Import<Assessment>(bytes, GroupImportOptions());

        Assert.Equal(2, imported.Count);
        Assert.Equal(90, imported[0].Performance);
        Assert.Equal(95, imported[0].Attendance);
        Assert.Equal(70, imported[1].Performance);
        Assert.Equal(80, imported[1].Attendance);
    }

    private static void SwapCells(ISheet sheet, int rowIndex, int colA, int colB)
    {
        var row = sheet.GetRow(rowIndex);
        var a = row.GetCell(colA);
        var b = row.GetCell(colB);
        string textA = a.ToString() ?? "";
        string textB = b.ToString() ?? "";
        a.SetCellValue(textB);
        b.SetCellValue(textA);
    }
}
