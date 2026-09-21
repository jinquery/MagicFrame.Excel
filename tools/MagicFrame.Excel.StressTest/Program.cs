using System.Diagnostics;
using MagicFrame.Excel.Engine;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;

namespace MagicFrame.Excel.StressTest;

/// <summary>
/// MagicFrame.Excel 压力测试：泛型集合 导出/解析 Excel 计时。
/// <para>
/// 默认 5000 行 × 30 列，先预热一次（JIT / 表达式编译缓存），再跑 N 次取平均。
/// 与常规单元测试完全分开（独立可执行项目）。
/// </para>
/// 用法：
///   dotnet run --project tools/MagicFrame.Excel.StressTest
///   dotnet run --project tools/MagicFrame.Excel.StressTest -- 10000 30 3
///   dotnet run --project tools/MagicFrame.Excel.StressTest -- formula 5000 30
///   （默认：行数 列数 次数，列数上限为实体属性数 30；formula：含公式列的 预计算 vs 打开时重算 对比）
/// </summary>
public static class Program
{
    public static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        int totalRows = 8000;
        int totalCols = 40;

        if (args.Length > 0 && args[0].ToLowerInvariant() == "formula")
        {
            FormulaBenchmark(Parse(args, 1, totalRows), Math.Min(Parse(args, 2, totalCols), totalCols));
            return;
        }

        int rowCount = Parse(args, 0, totalRows);
        int colCount = Math.Min(Parse(args, 1, totalCols), totalCols);
        int iterations = Parse(args, 2, 3);

        Console.WriteLine("MagicFrame.Excel 压力测试");
        Console.WriteLine($"数据规模：{rowCount} 行 × {colCount + 1} 列（{colCount} 普通 + 1 值映射列），迭代 {iterations} 次");
        Console.WriteLine("==================================================");
        Console.WriteLine();

        // 程序化列定义：与实体 P01..P30 一一对应
        var columns = Enumerable.Range(1, colCount)
            .Select(i => new ExcelColumn { Name = $"列{i:D2}", Field = $"P{i:D2}", Order = i })
            .ToList();

        // 追加 1 列值映射（0/1 -> 女/男，unknown 哨兵兜底），压测映射导出/导入路径
        columns.Add(new ExcelColumn
        {
            Name = "性别",
            Field = "Gender",
            Order = colCount + 1,
            ValueMap = new Dictionary<object, string>
            {
                [0] = "女",
                [1] = "男",
                [MagicFrame.Excel.Converters.ValueMapper.UnknownKey] = "2",
            },
        });
        int effectiveCols = columns.Count;

        var engine = ExcelEngine.CreateDefault();

        // 预热：触发 JIT、表达式编译缓存（不参与统计）
        RunOnce(engine, BuildRows(rowCount), columns, rowCount, "预热", record: false);

        Console.WriteLine("  迭代    导出     导入     Managed     WorkingSet");
        double exportTotal = 0, importTotal = 0, sizeTotal = 0;
        for (int i = 1; i <= iterations; i++)
        {
            var rows = BuildRows(rowCount);
            var (exportMs, importMs, sizeBytes) = RunOnce(engine, rows, columns, rowCount, $"第 {i} 次", record: true);
            exportTotal += exportMs;
            importTotal += importMs;
            sizeTotal += sizeBytes;

            Console.WriteLine($"         {exportMs,6:F0} ms  {importMs,6:F0} ms  {GC.GetTotalMemory(false) / 1024.0 / 1024.0,8:F1} MB  {Environment.WorkingSet / 1024.0 / 1024.0,9:F1} MB");
        }

        // 强制 GC 后看基线：若回落到低位说明是瞬时峰值，否则存在未释放
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        Console.WriteLine($"GC 后   {GC.GetTotalMemory(false) / 1024.0 / 1024.0,13:F1} MB  {Environment.WorkingSet / 1024.0 / 1024.0,18:F1} MB");

        Console.WriteLine("==================================================");
        Console.WriteLine($"导出平均：{exportTotal / iterations / 1000.0:F3} s / 次（{(rowCount / Math.Max(exportTotal / iterations / 1000.0, 1e-9)):N0} 行/s）");
        Console.WriteLine($"导入平均：{importTotal / iterations / 1000.0:F3} s / 次（{(rowCount / Math.Max(importTotal / iterations / 1000.0, 1e-9)):N0} 行/s）");
        Console.WriteLine($"文件平均：{sizeTotal / iterations / 1024.0 / 1024.0:F2} MB");
        Console.ReadLine();
    }

    /// <summary>含公式列的导出耗时对比：预计算（EvaluateAllFormulaCells）vs 打开时重算（fullCalcOnLoad）</summary>
    private static void FormulaBenchmark(int rowCount, int colCount)
    {
        Console.WriteLine("MagicFrame.Excel 公式导出压测");
        Console.WriteLine($"数据规模：{rowCount} 行 × {colCount + 1} 列（含 1 个公式列：P11*P12），各 1 次");
        Console.WriteLine("==================================================");
        Console.WriteLine();

        var columns = Enumerable.Range(1, colCount)
            .Select(i => new ExcelColumn { Name = $"列{i:D2}", Field = $"P{i:D2}", Order = i })
            .ToList();
        columns.Add(new ExcelColumn { Name = "合计", Field = "Total", Order = colCount + 1, Formula = "{P11}*{P12}" });

        var engine = ExcelEngine.CreateDefault();
        var rows = BuildRows(rowCount);

        // 预热
        engine.ExportBytes(rows, new ExcelExportOptions { SheetName = "压力表", Columns = columns });

        // A. 预计算（默认）
        var sw = Stopwatch.StartNew();
        byte[] eager = engine.ExportBytes(rows, new ExcelExportOptions { SheetName = "压力表", Columns = columns });
        sw.Stop();
        double eagerExport = sw.Elapsed.TotalMilliseconds;

        // B. 打开时重算（跳过 NPOI 求值）
        sw.Restart();
        byte[] deferred = engine.ExportBytes(rows, new ExcelExportOptions
        {
            SheetName = "压力表",
            Columns = columns,
            CalculateFormulasOnExport = false,
        });
        sw.Stop();
        double deferredExport = sw.Elapsed.TotalMilliseconds;

        // 两种文件的导入耗时（导入均按公式求值）
        sw.Restart();
        var eagerImport = engine.Import<StressRow>(eager, new ExcelImportOptions { SheetName = "压力表", Columns = columns });
        double eagerImportMs = sw.Elapsed.TotalMilliseconds;
        sw.Restart();
        var deferredImport = engine.Import<StressRow>(deferred, new ExcelImportOptions { SheetName = "压力表", Columns = columns });
        double deferredImportMs = sw.Elapsed.TotalMilliseconds;

        // 正确性抽查（用非零行，避免 0*0 恒等掩盖错误）
        if (eagerImport.Count != rowCount || deferredImport.Count != rowCount)
        {
            throw new InvalidOperationException("导入行数不符");
        }
        decimal expectedTotal = (decimal)(rows[1].P11 * rows[1].P12);
        if (eagerImport[1].Total != expectedTotal || deferredImport[1].Total != expectedTotal)
        {
            throw new InvalidOperationException($"公式结果校验失败：期望 {expectedTotal}");
        }

        Console.WriteLine($"预计算（默认）   导出 {eagerExport,8:F0} ms | 导入 {eagerImportMs,8:F0} ms | 文件 {eager.Length / 1024.0 / 1024.0,6:F2} MB");
        Console.WriteLine($"打开时重算       导出 {deferredExport,8:F0} ms | 导入 {deferredImportMs,8:F0} ms | 文件 {deferred.Length / 1024.0 / 1024.0,6:F2} MB");
        Console.WriteLine($"导出提速：{(eagerExport - deferredExport) / 1000.0:F2} s（{eagerExport / Math.Max(deferredExport, 1e-9):F1}×）");
        Console.WriteLine("==================================================");
    }

    /// <summary>执行一次 导出 + 导入，校验正确性，返回耗时（毫秒）</summary>
    private static (double ExportMs, double ImportMs, long SizeBytes) RunOnce(
        ExcelEngine engine, List<StressRow> rows, List<ExcelColumn> columns, int rowCount, string label, bool record)
    {
        // 导出
        var sw = Stopwatch.StartNew();
        byte[] bytes = engine.ExportBytes(rows, new ExcelExportOptions { SheetName = "压力表", Columns = columns });
        sw.Stop();
        double exportMs = sw.Elapsed.TotalMilliseconds;

        // 导入
        sw.Restart();
        var imported = engine.Import<StressRow>(bytes, new ExcelImportOptions { SheetName = "压力表", Columns = columns });
        sw.Stop();
        double importMs = sw.Elapsed.TotalMilliseconds;

        // 正确性校验
        if (imported.Count != rowCount)
        {
            throw new InvalidOperationException($"导入行数不符：期望 {rowCount}，实际 {imported.Count}");
        }
        AssertEqual(rows[0].P01, imported[0].P01, "P01[0]");
        AssertEqual(rows[0].P11, imported[0].P11, "P11[0]");
        AssertEqual(rows[rowCount - 1].P21, imported[rowCount - 1].P21, "P21[末行]");
        AssertEqual(rows[0].Gender, imported[0].Gender, "Gender[0]（值映射往返）");
        AssertEqual(rows[rowCount / 2].P26, imported[rowCount / 2].P26, "P26[中行]");
        AssertEqual(rows[rowCount / 3].P29, imported[rowCount / 3].P29, "P29[中行]");

        if (record)
        {
            Console.WriteLine($"{label,-8} 导出 {exportMs,8:F0} ms | 导入 {importMs,8:F0} ms | 文件 {bytes.Length / 1024.0 / 1024.0,6:F2} MB | 校验通过");
        }
        return (exportMs, importMs, bytes.Length);
    }

    private static List<StressRow> BuildRows(int rowCount)
    {
        var rows = new List<StressRow>(rowCount);
        for (int i = 0; i < rowCount; i++)
        {
            rows.Add(StressRow.Create(i));
        }
        return rows;
    }

    private static void AssertEqual<T>(T expected, T actual, string name)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
        {
            throw new InvalidOperationException($"校验失败 {name}：期望 {expected}，实际 {actual}");
        }
    }

    private static int Parse(string[] args, int index, int defaultValue)
        => args.Length > index && int.TryParse(args[index], out var v) && v > 0 ? v : defaultValue;
}
