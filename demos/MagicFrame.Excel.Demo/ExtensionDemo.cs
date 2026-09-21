using MagicFrame.Excel.Abstractions;
using MagicFrame.Excel.Engine;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;

namespace MagicFrame.Excel.Demo;

/// <summary>
/// 扩展点演示：不改动核心引擎，通过 Builder 替换/追加扩展实现。
/// 覆盖：自定义列提供器、自定义访问器、自定义值转换器、自定义公式解析器、自定义表头策略。
/// </summary>
public static class ExtensionDemo
{
    public static void Run(string outputDir)
    {
        Directory.CreateDirectory(outputDir);
        Console.WriteLine("==================== 扩展点演示 ====================");
        Console.WriteLine();

        // 1. 自定义列提供器：实体无需任何特性
        Console.WriteLine("[1] 自定义列提供器（无特性实体 + 程序化列定义）→ manual_columns.xlsx");
        var provider = new ManualColumnProvider(new[]
        {
            new ExcelColumn { Name = "编号", Field = "Id", Order = 1, Width = 8 },
            new ExcelColumn { Name = "名称", Field = "Name", Order = 2, Width = 16 },
            new ExcelColumn { Name = "金额", Field = "Amount", Order = 3, Width = 14 },
        });
        var providerEngine = ExcelEngineBuilder.Create().UseColumnProvider(provider).Build();

        var manualRows = new[]
        {
            new ManualRow { Id = 1, Name = "苹果", Amount = 10.5m },
            new ManualRow { Id = 2, Name = "香蕉", Amount = 20.0m },
        };
        string manualFile = Path.Combine(outputDir, "manual_columns.xlsx");
        providerEngine.ExportToFile(manualRows, manualFile, new ExcelExportOptions { SheetName = "手工列" });
        Console.WriteLine($"    已生成: {manualFile}");
        Console.WriteLine();

        // 2. 自定义访问器：字典行（无需强类型实体）
        Console.WriteLine("[2] 自定义访问器（字典行导出+解析）→ dictionary.xlsx");
        var dictAccessor = new DictionaryAccessor();
        var dictEngine = ExcelEngineBuilder.Create().UseEntityAccessor(dictAccessor).Build();
        var dictColumns = new[]
        {
            new ExcelColumn { Name = "键", Field = "Key", Order = 1 },
            new ExcelColumn { Name = "值", Field = "Value", Order = 2 },
        };
        var dictRows = new[]
        {
            new Dictionary<string, object?> { ["Key"] = "k1", ["Value"] = "v1" },
            new Dictionary<string, object?> { ["Key"] = "k2", ["Value"] = "v2" },
        };
        string dictFile = Path.Combine(outputDir, "dictionary.xlsx");
        dictEngine.ExportToFile(dictRows, dictFile, new ExcelExportOptions { SheetName = "字典", Columns = dictColumns });
        var dictBack = dictEngine.ImportFile<Dictionary<string, object?>>(dictFile, new ExcelImportOptions { SheetName = "字典", Columns = dictColumns });
        Console.WriteLine($"    已生成: {dictFile}，解析回 {dictBack.Count} 行（k1={dictBack[0]["Value"]}）");
        Console.WriteLine();

        // 3. 自定义值转换器：货币格式化
        Console.WriteLine("[3] 自定义值转换器（金额格式化 ¥#,##0.00）→ currency.xlsx");
        var currencyEngine = ExcelEngineBuilder.Create().UseCellValueConverter(new CurrencyConverter()).Build();
        string currencyFile = Path.Combine(outputDir, "currency.xlsx");
        currencyEngine.ExportToFile(
            new[] { new CurrencyItem { Amount = 1234.5m } },
            currencyFile,
            new ExcelExportOptions
            {
                SheetName = "货币",
                Columns = new[] { new ExcelColumn { Name = "金额", Field = "Amount", Order = 1, WriteAsNumeric = false } },
            });
        Console.WriteLine($"    已生成: {currencyFile}");
        Console.WriteLine();

        // 4. 自定义公式解析器：改用 [[Field]] 占位符
        Console.WriteLine("[4] 自定义公式解析器（[[Field]] 占位符）→ custom_formula.xlsx");
        var formulaEngine = ExcelEngineBuilder.Create().UseFormulaResolver(new DoubleBracketFormulaResolver()).Build();
        var formulaColumns = new[]
        {
            new ExcelColumn { Name = "数量", Field = "Qty", Order = 1 },
            new ExcelColumn { Name = "单价", Field = "Price", Order = 2 },
            new ExcelColumn { Name = "合计", Field = "Total", Order = 3, Formula = "[[Qty]]*[[Price]]" },
        };
        string formulaFile = Path.Combine(outputDir, "custom_formula.xlsx");
        formulaEngine.ExportToFile(
            new[] { new FormulaRow { Qty = 4, Price = 5m } },
            formulaFile,
            new ExcelExportOptions { SheetName = "公式", Columns = formulaColumns });
        Console.WriteLine($"    已生成: {formulaFile}（合计列使用 =A2*B2）");
        Console.WriteLine();

        // 5. 自定义表头策略：注册新的 HeaderKind "FixedTitle"
        Console.WriteLine("[5] 自定义表头策略（FixedTitle：固定标题 + 列名）→ fixed_title.xlsx");
        var fixedEngine = ExcelEngineBuilder.Create()
            .RegisterSheetWriter("FixedTitle", new FixedTitleSheetWriter())
            .RegisterSheetReader("FixedTitle", new FixedTitleSheetReader())
            .Build();
        var fixedRows = new[] { new ManualRow { Id = 1, Name = "苹果", Amount = 10.5m } };
        var fixedColumns = new[]
        {
            new ExcelColumn { Name = "编号", Field = "Id", Order = 1 },
            new ExcelColumn { Name = "名称", Field = "Name", Order = 2 },
        };
        string fixedFile = Path.Combine(outputDir, "fixed_title.xlsx");
        fixedEngine.ExportToFile(fixedRows, fixedFile, new ExcelExportOptions
        {
            SheetName = "固定标题",
            HeaderKind = "FixedTitle",
            Columns = fixedColumns,
        });
        var fixedBack = fixedEngine.ImportFile<ManualRow>(fixedFile, new ExcelImportOptions
        {
            HeaderKind = "FixedTitle",
            DataStartRowIndex = 2,
            Columns = fixedColumns,
        });
        Console.WriteLine($"    已生成: {fixedFile}，解析回 {fixedBack.Count} 行（名称={fixedBack[0].Name}）");
        Console.WriteLine();

        Console.WriteLine("==================== 扩展演示结束 ====================");
        Console.WriteLine();
    }

    // ==================== 扩展实现 ====================

    public class CurrencyItem
    {
        public decimal Amount { get; set; }
    }

    public class FormulaRow
    {
        public int Qty { get; set; }
        public decimal Price { get; set; }
    }

    /// <summary>手工列提供器</summary>
    public sealed class ManualColumnProvider : IColumnProvider
    {
        private readonly IReadOnlyList<ExcelColumn> _columns;
        public ManualColumnProvider(IReadOnlyList<ExcelColumn> columns) => _columns = columns;
        public IReadOnlyList<ExcelColumn> GetColumns(Type entityType) => _columns;
    }

    /// <summary>字典访问器</summary>
    public sealed class DictionaryAccessor : IEntityAccessor
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

    /// <summary>货币格式化转换器</summary>
    public sealed class CurrencyConverter : ICellValueConverter
    {
        public object? ReadCell(ICell? cell, Type targetType)
        {
            var value = MagicFrame.Excel.Converters.DefaultCellValueConverter.Instance.ReadCell(cell, targetType);
            if (value is string s && s.StartsWith("¥"))
            {
                return decimal.Parse(s.TrimStart('¥'), System.Globalization.CultureInfo.InvariantCulture);
            }
            return value;
        }

        public string? ToCellText(object? value)
            => value is decimal d ? "¥" + d.ToString("N2") : value?.ToString();

        public bool IsNumeric(object? value) => false;
    }

    /// <summary>双花括号公式解析器</summary>
    public sealed class DoubleBracketFormulaResolver : IFormulaResolver
    {
        private static readonly System.Text.RegularExpressions.Regex Regex =
            new(@"\[\[([^{}\[\]]+)\]\]", System.Text.RegularExpressions.RegexOptions.Compiled);

        public string Resolve(string template, IReadOnlyDictionary<string, string> fieldToAddress)
            => Regex.Replace(template, m =>
                fieldToAddress.TryGetValue(m.Groups[1].Value, out string? address) ? address : m.Value);
    }

    /// <summary>自定义表头：第 1 行固定标题、第 2 行列名、数据从第 3 行开始</summary>
    public sealed class FixedTitleSheetWriter : ISheetWriter
    {
        public int HeaderRowCount => 2;

        public void Write(
            ISheet sheet, IWorkbook workbook, IReadOnlyList<ExcelColumn> columns,
            IEntityAccessor accessor, ICellValueConverter converter, IFormulaResolver formulaResolver,
            IEnumerable<object> rows, ExcelExportOptions options)
        {
            sheet.CreateRow(0).CreateCell(0).SetCellValue("MagicFrame.Excel 固定标题示例");
            var header = sheet.CreateRow(1);
            for (int c = 0; c < columns.Count; c++)
            {
                header.CreateCell(c).SetCellValue(columns[c].Name);
                sheet.SetColumnWidth(c, Math.Max(columns[c].Width ?? 10, 8) * 256);
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

    /// <summary>自定义表头读取：第 1 行固定标题跳过，第 2 行列名，数据从第 3 行开始</summary>
    public sealed class FixedTitleSheetReader : ISheetReader
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
                indexByHeader[header.GetCell(c)?.ToString()?.Trim() ?? ""] = c;
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
