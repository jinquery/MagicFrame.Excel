using MagicFrame.Excel.Engine;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace MagicFrame.Excel.Tests;

/// <summary>
/// 新增功能测试：自动筛选 / 冻结列 / 数字格式 / 数据验证 / 空密码保护 /
/// 逐行错误收集 / 源行号回填 / 多 Sheet 一键导入 / 进度回调 / 样式扩展 / 结构化问题 / 列定义缓存隔离
/// </summary>
public class FeatureTests
{
    private readonly ExcelEngine _engine = TestHelper.Engine;

    private static List<Employee> Employees() => TestHelper.SampleEmployees();

    // ---------- C1 自动筛选 ----------

    [Fact]
    public void AutoFilter_Creates_Filter_Range()
    {
        using var result = _engine.Export(Employees(), new ExcelExportOptions { SheetName = "员工表", AutoFilter = true });
        var sheet = Assert.IsType<XSSFSheet>(result.Workbook.GetSheetAt(0));
        Assert.NotNull(sheet.GetCTWorksheet().autoFilter);
        Assert.Equal("A1:G1", sheet.GetCTWorksheet().autoFilter.@ref);
    }

    [Fact]
    public void No_AutoFilter_By_Default()
    {
        using var result = _engine.Export(Employees(), new ExcelExportOptions { SheetName = "员工表" });
        var sheet = Assert.IsType<XSSFSheet>(result.Workbook.GetSheetAt(0));
        Assert.Null(sheet.GetCTWorksheet().autoFilter);
    }

    // ---------- C3 冻结列 ----------

    [Fact]
    public void Freeze_Columns_And_Header_Work_Together()
    {
        using var result = _engine.Export(Employees(), new ExcelExportOptions { SheetName = "员工表", FreezeColumns = 2 });
        var sheet = Assert.IsType<XSSFSheet>(result.Workbook.GetSheetAt(0));
        var pane = sheet.PaneInformation;
        Assert.NotNull(pane);
        Assert.Equal(2, pane.VerticalSplitLeftColumn); // 冻结前 2 列
        Assert.Equal(1, pane.HorizontalSplitTopRow);   // 表头行仍冻结
    }

    // ---------- C2 数字格式 ----------

    [Fact]
    public void NumberFormat_Applied_To_Cell()
    {
        var columns = new[] { new ExcelColumn { Name = "比率", Field = "Ratio", Order = 1, NumberFormat = "0.00%" } };
        var rows = new[] { new ValidationRow { Ratio = 0.125 } };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表", Columns = columns });
        var cell = result.Workbook.GetSheetAt(0).GetRow(1).GetCell(0);
        string fmt = result.Workbook.CreateDataFormat().GetFormat(cell.CellStyle.DataFormat);
        Assert.Equal("0.00%", fmt);
    }

    // ---------- C4 数据验证增强 ----------

    [Fact]
    public void Validation_Integer_And_Custom_Are_Created()
    {
        var columns = new[]
        {
            new ExcelColumn { Name = "数量", Field = "Qty", Order = 1, ValidationKind = ValidationKind.Integer, ValidationFormula1 = "1", ValidationFormula2 = "100" },
            new ExcelColumn { Name = "备注", Field = "Note", Order = 2, ValidationKind = ValidationKind.CustomFormula, ValidationFormula1 = "A1>0" },
        };
        var rows = new[] { new ValidationRow { Qty = 5, Note = "x" } };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表", Columns = columns });
        var validations = result.Workbook.GetSheetAt(0).GetDataValidations();

        Assert.NotEmpty(validations);

        var intV = validations.First(v => v.Regions.CountRanges() > 0 && v.Regions.GetCellRangeAddress(0).FirstColumn == 0);
        Assert.Equal("1", intV.ValidationConstraint.Formula1);
        Assert.Equal("100", intV.ValidationConstraint.Formula2);
        Assert.Equal(NPOI.SS.UserModel.OperatorType.BETWEEN, intV.ValidationConstraint.Operator);

        var customV = validations.First(v => v.Regions.GetCellRangeAddress(0).FirstColumn == 1);
        Assert.Equal("A1>0", customV.ValidationConstraint.Formula1);
    }

    // ---------- C5 空密码保护 ----------

    [Fact]
    public void Protect_Without_Password()
    {
        var rows = new List<Account> { new() { Login = "a", Password = "b", Name = "c" } };
        var options = new ExcelExportOptions
        {
            SheetName = "账号表",
            Protection = new SheetProtectionOptions { Password = "", ProtectWithoutPassword = true },
        };

        using var result = _engine.Export(rows, options);
        var sheet = Assert.IsType<XSSFSheet>(result.Workbook.GetSheetAt(0));
        Assert.True(sheet.Protect);
        Assert.True(string.IsNullOrEmpty(sheet.GetCTWorksheet().sheetProtection!.password)); // 未写密码哈希
    }

    [Fact]
    public void Empty_Password_Without_Flag_Means_No_Protection()
    {
        var rows = new List<Account> { new() { Login = "a", Password = "b", Name = "c" } };
        var options = new ExcelExportOptions
        {
            SheetName = "账号表",
            Protection = new SheetProtectionOptions { Password = "" },
        };

        using var result = _engine.Export(rows, options);
        Assert.False(result.Workbook.GetSheetAt(0).Protect);
    }

    // ---------- C6 逐行错误收集 ----------

    [Fact]
    public void Collect_Row_Errors_Records_And_Continues()
    {
        var rows = new List<ImportRow>
        {
            new() { Name = "A", Status = TestStatus.Active },
            new() { Name = "B", Status = TestStatus.Done },
        };
        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "导入表" });

        // 把第 2 条数据的"状态"单元格改成非法枚举字符串，触发转换异常
        var sheet = Assert.IsType<XSSFSheet>(result.Workbook.GetSheetAt(0));
        sheet.GetRow(2).GetCell(1).SetCellValue("INVALID_STATUS");
        using var ms = new MemoryStream();
        result.Workbook.Write(ms, true);

        var options = new ExcelImportOptions { SheetName = "导入表", CollectRowErrors = true };
        var imported = _engine.Import<ImportRow>(ms.ToArray(), options);

        // 非法行被跳过，其余行正常导入
        Assert.Single(imported);
        Assert.Equal("A", imported[0].Name);
        Assert.Contains(options.Issues, i => i.Code == "RowWriteFailed" && i.RowNumber == 3);
    }

    // ---------- C7 源行号回填 ----------

    [Fact]
    public void Source_Row_Property_Is_Backfilled()
    {
        var rows = new List<ImportRow>
        {
            new() { Name = "A", Status = TestStatus.Active },
            new() { Name = "B", Status = TestStatus.Done },
            new() { Name = "C", Status = TestStatus.Pending },
        };
        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "导入表" });

        var imported = _engine.Import<ImportRow>(result.Bytes, new ExcelImportOptions { SheetName = "导入表", SourceRowProperty = "SourceRow" });

        Assert.Equal(3, imported.Count);
        // 三个数据行对应 Excel 第 2/3/4 行
        Assert.Equal(2, imported[0].SourceRow);
        Assert.Equal(3, imported[1].SourceRow);
        Assert.Equal(4, imported[2].SourceRow);
    }

    // ---------- C8 多 Sheet 一键导入 ----------

    [Fact]
    public void ImportAll_Returns_All_Sheets()
    {
        var employees = Employees();
        var sheets = new Dictionary<string, IList<Employee>>
        {
            ["研发"] = employees.Where(e => e.Department == "研发").ToList(),
            ["测试"] = employees.Where(e => e.Department == "测试").ToList(),
        };
        using var result = _engine.ExportSheets(sheets);

        var all = _engine.ImportAll<Employee>(result.Bytes);

        Assert.Equal(2, all.Count);
        Assert.Single(all["研发"]);
        Assert.Equal("张三", all["研发"][0].Name);
        Assert.Single(all["测试"]);
        Assert.Equal("李四", all["测试"][0].Name);
    }

    // ---------- C9 进度回调 ----------

    private sealed class CaptureProgress : IProgress<double>
    {
        public List<double> Values { get; } = new();
        public void Report(double value) => Values.Add(value);
    }

    [Fact]
    public void Export_Progress_Is_Reported()
    {
        var progress = new CaptureProgress();
        var options = new ExcelExportOptions { SheetName = "员工表", Progress = progress };

        using var result = _engine.Export(Employees(), options);

        Assert.NotEmpty(progress.Values);
        Assert.Equal(1.0, progress.Values[^1], 4);
    }

    [Fact]
    public void Import_Progress_Is_Reported()
    {
        using var result = _engine.Export(Employees(), new ExcelExportOptions { SheetName = "员工表" });
        var progress = new CaptureProgress();
        var options = new ExcelImportOptions { SheetName = "员工表", Progress = progress };

        _engine.Import<Employee>(result.Bytes, options);

        Assert.NotEmpty(progress.Values);
        Assert.Equal(1.0, progress.Values[^1], 4);
    }

    // ---------- C10 样式扩展 ----------

    [Fact]
    public void Row_Height_Is_Applied()
    {
        using var result = _engine.Export(Employees(), new ExcelExportOptions { SheetName = "员工表", RowHeight = 30 });
        var row = result.Workbook.GetSheetAt(0).GetRow(1);
        Assert.Equal(30, row.HeightInPoints, 1);
    }

    [Fact]
    public void Borders_Are_Applied_To_Header_And_Data()
    {
        using var result = _engine.Export(Employees(), new ExcelExportOptions { SheetName = "员工表", EnableBorders = true });
        var sheet = result.Workbook.GetSheetAt(0);

        // 表头单元格有边框（整表网格）
        var headerCell = sheet.GetRow(0).GetCell(0);
        Assert.Equal(BorderStyle.Thin, headerCell.CellStyle.BorderTop);
        Assert.Equal(BorderStyle.Thin, headerCell.CellStyle.BorderLeft);
        // 数据单元格有边框
        var dataCell = sheet.GetRow(1).GetCell(0);
        Assert.Equal(BorderStyle.Thin, dataCell.CellStyle.BorderTop);
        Assert.Equal(BorderStyle.Thin, dataCell.CellStyle.BorderLeft);
    }

    [Fact]
    public void No_Borders_By_Default()
    {
        using var result = _engine.Export(Employees(), new ExcelExportOptions { SheetName = "员工表" });
        var cell = result.Workbook.GetSheetAt(0).GetRow(1).GetCell(0);
        Assert.Equal(BorderStyle.None, cell.CellStyle.BorderTop); // 默认无边框
    }

    [Fact]
    public void Alternate_Row_Fill_Is_Applied()
    {
        short color = NPOI.HSSF.Util.HSSFColor.LightBlue.Index;
        using var result = _engine.Export(Employees(), new ExcelExportOptions { SheetName = "员工表", AlternateRowFillColor = color });

        var sheet = result.Workbook.GetSheetAt(0);
        Assert.NotEqual(color, sheet.GetRow(1).GetCell(0).CellStyle.FillForegroundColor); // 第 1 数据行（偶数 r=0）不填充
        Assert.Equal(color, sheet.GetRow(2).GetCell(0).CellStyle.FillForegroundColor);    // 第 2 数据行（r=1）填充
    }

    // ---------- E4 结构化问题 ----------

    [Fact]
    public void Structured_Issue_For_Missing_Column()
    {
        using var result = _engine.Export(Employees(), new ExcelExportOptions { SheetName = "员工表" });
        var sheet = Assert.IsType<XSSFSheet>(result.Workbook.GetSheetAt(0));
        sheet.GetRow(0).GetCell(2).SetCellValue(""); // 清空"年龄"表头
        using var ms = new MemoryStream();
        result.Workbook.Write(ms, true);

        var options = new ExcelImportOptions { SheetName = "员工表" };
        Assert.Throws<Exceptions.ExcelImportException>(() => _engine.Import<Employee>(ms.ToArray(), options));

        Assert.Contains(options.Issues, i => i.Code == "MissingColumn" && i.Severity == ImportIssueSeverity.Error && i.RowNumber == null);
    }

    // ---------- A3/D1 列定义缓存隔离 ----------

    [Fact]
    public void Column_Provider_Returns_Isolated_Clones()
    {
        var provider = MagicFrame.Excel.Providers.AttributeColumnProvider.Instance;
        var first = provider.GetColumns(typeof(Employee));
        first[0].Name = "被篡改";

        var second = provider.GetColumns(typeof(Employee));
        Assert.Equal("姓名", second[0].Name); // 缓存未被污染
    }
}
