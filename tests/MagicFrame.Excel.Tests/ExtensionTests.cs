using MagicFrame.Excel.Abstractions;
using MagicFrame.Excel.Converters;
using MagicFrame.Excel.Engine;
using MagicFrame.Excel.Exceptions;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;
using Xunit;

namespace MagicFrame.Excel.Tests;

/// <summary>
/// 扩展点测试：验证"对修改关闭、对扩展开放"。
/// 通过 Builder 替换/追加 列提供器、访问器、转换器、公式解析器、表头策略，核心引擎代码零改动。
/// </summary>
public class ExtensionTests
{
    // ---------- 1. 自定义列提供器：手工构造列定义，实体无需任何特性 ----------

    [Fact]
    public void Custom_ColumnProvider_Without_Attributes()
    {
        var provider = new ManualColumnProvider(new[]
        {
            new ExcelColumn { Name = "编号", Field = "Id", Order = 1 },
            new ExcelColumn { Name = "名称", Field = "Name", Order = 2 },
        });
        var engine = ExcelEngineBuilder.Create().UseColumnProvider(provider).Build();

        var rows = new[] { new SimpleRow { Id = 1, Name = "A" }, new SimpleRow { Id = 2, Name = "B" } };
        var bytes = engine.ExportBytes(rows);

        var grid = TestHelper.DumpSheet(new NPOI.XSSF.UserModel.XSSFWorkbook(new MemoryStream(bytes)));
        Assert.Equal("编号", grid[0, 0]);
        Assert.Equal("名称", grid[0, 1]);
        Assert.Equal("1", grid[1, 0]);
        Assert.Equal("A", grid[1, 1]);
    }

    // ---------- 2. 自定义访问器：字典行（不依赖强类型实体） ----------

    [Fact]
    public void Custom_Accessor_For_Dictionary_Rows()
    {
        var engine = ExcelEngineBuilder.Create().UseEntityAccessor(new DictionaryAccessor()).Build();
        var columns = new[]
        {
            new ExcelColumn { Name = "键", Field = "Key", Order = 1 },
            new ExcelColumn { Name = "值", Field = "Value", Order = 2 },
        };
        var options = new ExcelExportOptions { Columns = columns };
        var rows = new[] { new Dictionary<string, object?> { ["Key"] = "k1", ["Value"] = "v1" } };

        var bytes = engine.ExportBytes(rows, options);
        var imported = engine.Import<Dictionary<string, object?>>(bytes, new ExcelImportOptions { Columns = columns });

        Assert.Single(imported);
        Assert.Equal("k1", imported[0]["Key"]);
        Assert.Equal("v1", imported[0]["Value"]);
    }

    // ---------- 3. 自定义值转换器：货币格式化 ----------

    [Fact]
    public void Custom_Converter_Formats_Values()
    {
        var engine = ExcelEngineBuilder.Create().UseCellValueConverter(new CurrencyConverter()).Build();
        var rows = new[] { new MoneyItem { Amount = 1234.5m } };
        var columns = new[] { new ExcelColumn { Name = "金额", Field = "Amount", Order = 1, WriteAsNumeric = false } };

        var bytes = engine.ExportBytes(rows, new ExcelExportOptions { Columns = columns });
        using var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(new MemoryStream(bytes));
        var cell = workbook.GetSheetAt(0).GetRow(1).GetCell(0);

        Assert.Equal("¥1,234.50", cell.StringCellValue);
    }

    // ---------- 4. 自定义公式解析器：改用 [[Field]] 占位符 ----------

    [Fact]
    public void Custom_FormulaResolver_With_Different_Syntax()
    {
        var engine = ExcelEngineBuilder.Create().UseFormulaResolver(new DoubleBracketFormulaResolver()).Build();
        var columns = new[]
        {
            new ExcelColumn { Name = "数量", Field = "Qty", Order = 1 },
            new ExcelColumn { Name = "单价", Field = "Price", Order = 2 },
            new ExcelColumn { Name = "合计", Field = "Total", Order = 3, Formula = "[[Qty]]*[[Price]]" },
        };
        var rows = new[] { new FormulaRow { Qty = 4, Price = 5.0, Total = 0.0 } };

        var bytes = engine.ExportBytes(rows, new ExcelExportOptions { Columns = columns });
        using var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(new MemoryStream(bytes));
        var formulaCell = workbook.GetSheetAt(0).GetRow(1).GetCell(2);

        Assert.Equal(CellType.Formula, formulaCell.CellType);
        Assert.Equal("A2*B2", formulaCell.CellFormula);
        Assert.Equal(20.0, formulaCell.NumericCellValue, 4);
    }

    // ---------- 5. 自定义表头策略：注册新的 HeaderKind ----------

    [Fact]
    public void Custom_HeaderKind_Registered_To_Builder()
    {
        var engine = ExcelEngineBuilder.Create()
            .RegisterSheetWriter("FixedTitle", new FixedTitleSheetWriter())
            .RegisterSheetReader("FixedTitle", new FixedTitleSheetReader())
            .Build();

        var rows = new List<TitleItem> { new() { Col1 = "x", Col2 = "y" } };
        var exportOptions = new ExcelExportOptions { SheetName = "表", HeaderKind = "FixedTitle" };
        var bytes = engine.ExportBytes(rows, exportOptions);

        using var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(new MemoryStream(bytes));
        var sheet = workbook.GetSheetAt(0);
        Assert.Equal("这是一行固定标题", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("标题列1", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("x", sheet.GetRow(2).GetCell(0).StringCellValue);

        var imported = engine.Import<TitleItem>(bytes, new ExcelImportOptions { HeaderKind = "FixedTitle", DataStartRowIndex = 2 });
        Assert.Single(imported);
        Assert.Equal("x", imported[0].Col1);
        Assert.Equal("y", imported[0].Col2);
    }

    [Fact]
    public void Unregistered_HeaderKind_Throws()
    {
        var engine = ExcelEngine.CreateDefault();
        var ex = Assert.Throws<ExcelImportException>(() =>
            engine.ExportBytes(TestHelper.SampleEmployees(), new ExcelExportOptions { HeaderKind = "NotExist" }));
        Assert.Contains("未注册", ex.Message);
    }

    // ==================== 辅助类型 ====================

    public class MoneyItem
    {
        public decimal Amount { get; set; }
    }

    /// <summary>手工列提供器测试用的普通类（无任何特性）</summary>
    public class SimpleRow
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
    }

    /// <summary>自定义公式解析器测试用的普通类（无任何特性）</summary>
    public class FormulaRow
    {
        public int Qty { get; set; }
        public double Price { get; set; }
        public double Total { get; set; }
    }

    /// <summary>手工列提供器</summary>
    private sealed class ManualColumnProvider : IColumnProvider
    {
        private readonly IReadOnlyList<ExcelColumn> _columns;
        public ManualColumnProvider(IReadOnlyList<ExcelColumn> columns) => _columns = columns;
        public IReadOnlyList<ExcelColumn> GetColumns(Type entityType) => _columns;
    }

    /// <summary>字典访问器：支持 Dictionary[string, object?] 行</summary>
    private sealed class DictionaryAccessor : IEntityAccessor
    {
        public object? Read(object entity, string field)
            => entity is IDictionary<string, object?> dict && dict.TryGetValue(field, out var v) ? v : null;

        public void Write(object entity, string field, object? value)
        {
            if (entity is IDictionary<string, object?> dict)
            {
                dict[field] = value;
            }
        }
    }

    /// <summary>货币格式化转换器：把 decimal 渲染为 ¥#,###.##</summary>
    private sealed class CurrencyConverter : ICellValueConverter
    {
        public object? ReadCell(ICell? cell, Type targetType)
        {
            var value = DefaultCellValueConverter.Instance.ReadCell(cell, targetType);
            if (value is string s && s.StartsWith("¥"))
            {
                return decimal.Parse(s.TrimStart('¥'), System.Globalization.CultureInfo.InvariantCulture);
            }
            return value;
        }

        public string? ToCellText(object? value)
            => value is decimal d ? "¥" + d.ToString("N2") : value?.ToString();

        public bool IsNumeric(object? value) => false; // 一律按文本写入
    }

    /// <summary>双花括号公式解析器：把 [[Field]] 替换为单元格地址</summary>
    private sealed class DoubleBracketFormulaResolver : IFormulaResolver
    {
        private static readonly System.Text.RegularExpressions.Regex Regex =
            new(@"\[\[([^{}\[\]]+)\]\]", System.Text.RegularExpressions.RegexOptions.Compiled);

        public string Resolve(string template, IReadOnlyDictionary<string, string> fieldToAddress)
            => Regex.Replace(template, m =>
                fieldToAddress.TryGetValue(m.Groups[1].Value, out string? address) ? address : m.Value);
    }

    /// <summary>自定义表头策略：第 1 行固定标题，第 2 行列名，数据从第 3 行开始</summary>
    private sealed class FixedTitleSheetWriter : ISheetWriter
    {
        public int HeaderRowCount => 2;

        public void Write(
            ISheet sheet, IWorkbook workbook, IReadOnlyList<ExcelColumn> columns,
            IEntityAccessor accessor, ICellValueConverter converter, IFormulaResolver formulaResolver,
            IEnumerable<object> rows, ExcelExportOptions options)
        {
            sheet.CreateRow(0).CreateCell(0).SetCellValue("这是一行固定标题");
            var header = sheet.CreateRow(1);
            for (int c = 0; c < columns.Count; c++)
            {
                header.CreateCell(c).SetCellValue(columns[c].Name);
            }
            int r = 2;
            foreach (var row in rows)
            {
                var data = sheet.CreateRow(r++);
                for (int c = 0; c < columns.Count; c++)
                {
                    data.CreateCell(c).SetCellValue(accessor.Read(row, columns[c].Field)?.ToString() ?? "");
                }
            }
            sheet.CreateFreezePane(0, 2);
        }
    }

    /// <summary>自定义表头读取策略：第 1 行固定标题跳过，第 2 行列名，数据从第 3 行开始</summary>
    private sealed class FixedTitleSheetReader : ISheetReader
    {
        public IReadOnlyList<object> Read(
            ISheet sheet, IReadOnlyList<ExcelColumn> columns, Type entityType,
            IEntityAccessor accessor, ICellValueConverter converter, ExcelImportOptions options)
        {
            var result = new List<object>();
            var header = sheet.GetRow(1);
            var indexByHeader = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 0; c < header.LastCellNum; c++)
            {
                indexByHeader[header.GetCell(c).StringCellValue.Trim()] = c;
            }
            int start = options.DataStartRowIndex >= 3 ? options.DataStartRowIndex : 2;
            for (int r = start; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null)
                {
                    continue;
                }
                var entity = Activator.CreateInstance(entityType)!;
                foreach (var col in columns)
                {
                    if (indexByHeader.TryGetValue(col.Name, out int idx))
                    {
                        accessor.Write(entity, col.Field, converter.ReadCell(row.GetCell(idx), typeof(string)));
                    }
                }
                result.Add(entity);
            }
            return result;
        }
    }
}
