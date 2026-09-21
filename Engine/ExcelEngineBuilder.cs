using MagicFrame.Excel.Abstractions;
using MagicFrame.Excel.Accessors;
using MagicFrame.Excel.Converters;
using MagicFrame.Excel.Formulas;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Providers;
using MagicFrame.Excel.Readers;
using MagicFrame.Excel.Writers;

namespace MagicFrame.Excel.Engine;

/// <summary>
/// 引擎构建器：开放的扩展注册入口。
/// 所有扩展点都有默认实现，按需替换/追加即可，核心引擎代码零改动：
/// <list type="bullet">
/// <item>列元数据来源：<see cref="IColumnProvider"/>（默认特性发现）</item>
/// <item>实体读写：<see cref="IEntityAccessor"/>（默认反射）</item>
/// <item>单元格值转换：<see cref="ICellValueConverter"/>（默认通用转换）</item>
/// <item>公式占位符：<see cref="IFormulaResolver"/>（默认花括号替换）</item>
/// <item>表头布局：<see cref="ISheetWriter"/> + <see cref="ISheetReader"/>（默认单表头/分组表头）</item>
/// </list>
/// </summary>
public class ExcelEngineBuilder
{
    private IColumnProvider _columnProvider = AttributeColumnProvider.Instance;
    private IEntityAccessor _accessor = ReflectionEntityAccessor.Instance;
    private ICellValueConverter _converter = DefaultCellValueConverter.Instance;
    private IFormulaResolver _formulaResolver = DefaultFormulaResolver.Instance;
    private readonly SheetWriterRegistry _writers = new();
    private readonly SheetReaderRegistry _readers = new();

    private ExcelEngineBuilder()
    {
        RegisterDefaults();
    }

    /// <summary>创建构建器</summary>
    public static ExcelEngineBuilder Create() => new();

    private void RegisterDefaults()
    {
        _writers.Register(HeaderKinds.Single, new SingleHeaderSheetWriter());
        _writers.Register(HeaderKinds.Group, new GroupHeaderSheetWriter());
        _readers.Register(HeaderKinds.Single, new SingleHeaderSheetReader());
        _readers.Register(HeaderKinds.Group, new GroupHeaderSheetReader());
    }

    /// <summary>替换列元数据提供器</summary>
    public ExcelEngineBuilder UseColumnProvider(IColumnProvider provider)
    {
        _columnProvider = provider ?? throw new ArgumentNullException(nameof(provider));
        return this;
    }

    /// <summary>替换实体访问器</summary>
    public ExcelEngineBuilder UseEntityAccessor(IEntityAccessor accessor)
    {
        _accessor = accessor ?? throw new ArgumentNullException(nameof(accessor));
        return this;
    }

    /// <summary>替换单元格值转换器</summary>
    public ExcelEngineBuilder UseCellValueConverter(ICellValueConverter converter)
    {
        _converter = converter ?? throw new ArgumentNullException(nameof(converter));
        return this;
    }

    /// <summary>替换公式解析器</summary>
    public ExcelEngineBuilder UseFormulaResolver(IFormulaResolver resolver)
    {
        _formulaResolver = resolver ?? throw new ArgumentNullException(nameof(resolver));
        return this;
    }

    /// <summary>注册/覆盖指定表头类型的写入策略</summary>
    public ExcelEngineBuilder RegisterSheetWriter(string headerKind, ISheetWriter writer)
    {
        _writers.Register(headerKind, writer);
        return this;
    }

    /// <summary>注册/覆盖指定表头类型的读取策略</summary>
    public ExcelEngineBuilder RegisterSheetReader(string headerKind, ISheetReader reader)
    {
        _readers.Register(headerKind, reader);
        return this;
    }

    /// <summary>构建引擎</summary>
    public ExcelEngine Build()
        => new(_columnProvider, _accessor, _converter, _formulaResolver, _writers, _readers);
}
