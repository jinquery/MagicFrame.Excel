using MagicFrame.Excel.Abstractions;
using MagicFrame.Excel.Exceptions;

namespace MagicFrame.Excel.Writers;

/// <summary>
/// Sheet 写入策略注册表（开放扩展点）：
/// 新增表头布局时实现 <see cref="ISheetWriter"/> 并注册到新的 HeaderKind，
/// 引擎无需任何改动即可支持。
/// </summary>
public class SheetWriterRegistry
{
    private readonly Dictionary<string, ISheetWriter> _items = new(StringComparer.OrdinalIgnoreCase);

    public void Register(string headerKind, ISheetWriter writer)
    {
        if (string.IsNullOrWhiteSpace(headerKind))
        {
            throw new ArgumentException("headerKind 不能为空", nameof(headerKind));
        }
        _items[headerKind] = writer ?? throw new ArgumentNullException(nameof(writer));
    }

    public ISheetWriter Get(string headerKind)
    {
        if (_items.TryGetValue(headerKind, out var writer))
        {
            return writer;
        }
        throw new ExcelImportException($"未注册的 Sheet 写入策略：{headerKind}");
    }

    public bool Contains(string headerKind) => _items.ContainsKey(headerKind);
}
