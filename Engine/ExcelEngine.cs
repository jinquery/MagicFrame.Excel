using MagicFrame.Excel.Abstractions;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using MagicFrame.Excel.Readers;
using MagicFrame.Excel.Writers;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace MagicFrame.Excel.Engine;

/// <summary>
/// Excel 引擎门面：统一封装"泛型集合 生成 Excel / 解析 Excel"。
/// 该引擎对修改关闭（固定流水线：解析列定义 → 选择写入/读取策略 → 生成/解析），
/// 对扩展开放（通过 <see cref="ExcelEngineBuilder"/> 注册新的列提供器/访问器/转换器/公式解析器/表头策略）。
/// </summary>
public class ExcelEngine
{
    private readonly IColumnProvider _columnProvider;
    private readonly IEntityAccessor _accessor;
    private readonly ICellValueConverter _converter;
    private readonly IFormulaResolver _formulaResolver;
    private readonly SheetWriterRegistry _writers;
    private readonly SheetReaderRegistry _readers;

    internal ExcelEngine(
        IColumnProvider columnProvider,
        IEntityAccessor accessor,
        ICellValueConverter converter,
        IFormulaResolver formulaResolver,
        SheetWriterRegistry writers,
        SheetReaderRegistry readers)
    {
        _columnProvider = columnProvider;
        _accessor = accessor;
        _converter = converter;
        _formulaResolver = formulaResolver;
        _writers = writers;
        _readers = readers;
    }

    /// <summary>创建使用默认实现（特性列、反射访问器、默认转换/公式、单表头+分组表头）的引擎</summary>
    public static ExcelEngine CreateDefault() => ExcelEngineBuilder.Create().Build();

    // ==================== 导出 ====================

    /// <summary>泛型集合导出为 Excel 工作簿（返回字节/Base64/流/工作簿）</summary>
    public ExcelExportResult Export<T>(IEnumerable<T> rows, ExcelExportOptions? options = null) where T : new()
    {
        var opts = options ?? new ExcelExportOptions();
        opts.KnownRowCount ??= rows switch
        {
            ICollection<T> c => c.Count,
            IReadOnlyCollection<T> r => r.Count,
            _ => null,
        };
        var columns = ResolveColumns(typeof(T), opts.Columns);
        var workbook = new XSSFWorkbook();
        ISheet sheet = workbook.CreateSheet(opts.SheetName);
        var writer = _writers.Get(opts.HeaderKind);
        writer.Write(sheet, workbook, columns, _accessor, _converter, _formulaResolver, rows.Cast<object>(), opts);
        return ToResult(workbook);
    }

    /// <summary>导出为字节数组。注：工作簿内存由 GC 回收（非泄漏），不在此处 Dispose 以免拖慢导出；大文件如需主动释放请用 <see cref="Export{T}"/> 后自行 Dispose。</summary>
    public byte[] ExportBytes<T>(IEnumerable<T> rows, ExcelExportOptions? options = null) where T : new()
        => Export(rows, options).Bytes;

    /// <summary>导出为 Base64 字符串（供前端下载）</summary>
    public string ExportBase64<T>(IEnumerable<T> rows, ExcelExportOptions? options = null) where T : new()
        => Export(rows, options).Base64;

    /// <summary>导出为内存流</summary>
    public MemoryStream ExportStream<T>(IEnumerable<T> rows, ExcelExportOptions? options = null) where T : new()
        => Export(rows, options).Stream;

    /// <summary>导出到本地文件</summary>
    public void ExportToFile<T>(IEnumerable<T> rows, string filePath, ExcelExportOptions? options = null) where T : new()
        => File.WriteAllBytes(filePath, ExportBytes(rows, options));

    /// <summary>多 Sheet 导出</summary>
    public ExcelExportResult ExportSheets(IEnumerable<SheetSpec> sheets, ExcelExportOptions? defaultOptions = null)
    {
        var workbook = new XSSFWorkbook();
        foreach (var spec in sheets)
        {
            if (spec.RowType == null)
            {
                throw new ArgumentException("SheetSpec.RowType 不能为空", nameof(sheets));
            }
            var opts = spec.Options ?? defaultOptions ?? new ExcelExportOptions();
            if (!string.IsNullOrWhiteSpace(spec.Name))
            {
                opts.SheetName = spec.Name;
            }
            opts.KnownRowCount ??= spec.Rows switch
            {
                ICollection<object> c => c.Count,
                IReadOnlyCollection<object> r => r.Count,
                _ => null,
            };
            var columns = ResolveColumns(spec.RowType, opts.Columns);
            ISheet sheet = workbook.CreateSheet(opts.SheetName);
            var writer = _writers.Get(opts.HeaderKind);
            writer.Write(sheet, workbook, columns, _accessor, _converter, _formulaResolver, spec.Rows, opts);
        }
        return ToResult(workbook);
    }

    /// <summary>多 Sheet 导出（同类型便捷重载）</summary>
    public ExcelExportResult ExportSheets<T>(IDictionary<string, IList<T>> sheets, ExcelExportOptions? defaultOptions = null) where T : new()
        => ExportSheets(sheets.Select(s => SheetSpec.Of(s.Key, s.Value, null)), defaultOptions);

    // ==================== 导入 ====================

    /// <summary>从字节数组解析 Excel 到泛型集合</summary>
    public IList<T> Import<T>(byte[] bytes, ExcelImportOptions? options = null) where T : new()
    {
        using var stream = new MemoryStream(bytes);
        return Import<T>(stream, options);
    }

    /// <summary>从 Base64 字符串解析 Excel 到泛型集合</summary>
    public IList<T> Import<T>(string base64, ExcelImportOptions? options = null) where T : new()
        => Import<T>(Convert.FromBase64String(base64), options);

    /// <summary>从文件解析 Excel 到泛型集合</summary>
    public IList<T> ImportFile<T>(string filePath, ExcelImportOptions? options = null) where T : new()
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        return Import<T>(stream, options);
    }

    /// <summary>从流解析 Excel 到泛型集合</summary>
    public IList<T> Import<T>(Stream stream, ExcelImportOptions? options = null) where T : new()
    {
        var opts = options ?? new ExcelImportOptions();
        using var workbook = new XSSFWorkbook(stream);
        ISheet sheet = ResolveSheet(workbook, opts);
        var columns = ResolveColumns(typeof(T), opts.Columns);
        var reader = _readers.Get(opts.HeaderKind);
        var list = reader.Read(sheet, columns, typeof(T), _accessor, _converter, opts);
        return list.Cast<T>().ToList();
    }

    /// <summary>把工作簿所有 Sheet 按同一实体类型解析（C8）：返回 Sheet 名 -> 实体集合</summary>
    public IDictionary<string, IList<T>> ImportAll<T>(byte[] bytes, ExcelImportOptions? options = null) where T : new()
    {
        var result = new Dictionary<string, IList<T>>(StringComparer.OrdinalIgnoreCase);
        using var stream = new MemoryStream(bytes);
        using var workbook = new XSSFWorkbook(stream);
        for (int i = 0; i < workbook.NumberOfSheets; i++)
        {
            var sheet = workbook.GetSheetAt(i);
            var opts = (options ?? new ExcelImportOptions()).CopyForSheet(sheet.SheetName);
            var columns = ResolveColumns(typeof(T), opts.Columns);
            var reader = _readers.Get(opts.HeaderKind);
            var list = reader.Read(sheet, columns, typeof(T), _accessor, _converter, opts);
            result[sheet.SheetName] = list.Cast<T>().ToList();
        }
        return result;
    }

    // ==================== 内部 ====================

    private IReadOnlyList<ExcelColumn> ResolveColumns(Type entityType, IReadOnlyList<ExcelColumn>? configured)
    {
        if (configured is { Count: > 0 })
        {
            return configured;
        }
        var columns = _columnProvider.GetColumns(entityType);
        if (columns.Count == 0)
        {
            throw new InvalidOperationException(
                $"实体类型 {entityType.FullName} 未配置任何列：请添加 {nameof(Attributes.ExcelColumnAttribute)} 特性或通过选项传入列定义。");
        }
        return columns;
    }

    private static ISheet ResolveSheet(IWorkbook workbook, ExcelImportOptions options)
    {
        if (!string.IsNullOrWhiteSpace(options.SheetName))
        {
            return workbook.GetSheet(options.SheetName)
                   ?? throw new Exceptions.ExcelImportException($"找不到 Sheet：{options.SheetName}");
        }
        return workbook.GetSheetAt(options.SheetIndex);
    }

    private static ExcelExportResult ToResult(XSSFWorkbook workbook)
    {
        var stream = new MemoryStream();
        workbook.Write(stream, true); // leaveOpen：workbook.Write 默认会关闭流
        stream.Position = 0;
        var bytes = stream.ToArray();
        return new ExcelExportResult(workbook, stream, bytes);
    }
}
