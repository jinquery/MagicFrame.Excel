using MagicFrame.Excel.Abstractions;
using MagicFrame.Excel.Exceptions;

namespace MagicFrame.Excel.Readers;

/// <summary>
/// Sheet 读取策略注册表（开放扩展点）：
/// 新增表头布局时实现 <see cref="ISheetReader"/> 并注册到对应 HeaderKind。
/// </summary>
public class SheetReaderRegistry
{
    private readonly Dictionary<string, ISheetReader> _items = new(StringComparer.OrdinalIgnoreCase);

    public void Register(string headerKind, ISheetReader reader)
    {
        if (string.IsNullOrWhiteSpace(headerKind))
        {
            throw new ArgumentException("headerKind 不能为空", nameof(headerKind));
        }
        _items[headerKind] = reader ?? throw new ArgumentNullException(nameof(reader));
    }

    public ISheetReader Get(string headerKind)
    {
        if (_items.TryGetValue(headerKind, out var reader))
        {
            return reader;
        }
        throw new ExcelImportException($"未注册的 Sheet 读取策略：{headerKind}");
    }

    public bool Contains(string headerKind) => _items.ContainsKey(headerKind);
}
