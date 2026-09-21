using System.Collections.Concurrent;
using MagicFrame.Excel.Model;
using NPOI.SS.UserModel;

namespace MagicFrame.Excel.Writers;

/// <summary>
/// 单次写入使用的样式集合：每个 Workbook 独立创建，避免跨工作簿复用样式导致异常。
/// 因此写入策略（ISheetWriter 实现）可安全地作为单例复用。
/// 数据样式按 (锁定/日期/交替行/边框/数字格式) 组合缓存，避免为每列/每行重复建样式。
/// </summary>
public sealed class WriterStyles
{
    private readonly IWorkbook _workbook;
    private readonly Dictionary<short, ICellStyle> _headersByColor = new();
    private readonly ICellStyle _cellStyle;
    private readonly ICellStyle _lockedStyle;
    private readonly ICellStyle _dateCellStyle;
    private readonly ICellStyle _dateLockedStyle;
    private readonly ConcurrentDictionary<string, ICellStyle> _dataStyles = new();

    private WriterStyles(
        IWorkbook workbook,
        ICellStyle cellStyle,
        ICellStyle lockedStyle,
        ICellStyle dateCellStyle,
        ICellStyle dateLockedStyle)
    {
        _workbook = workbook;
        _cellStyle = cellStyle;
        _lockedStyle = lockedStyle;
        _dateCellStyle = dateCellStyle;
        _dateLockedStyle = dateLockedStyle;
    }

    /// <summary>为指定工作簿创建样式集合</summary>
    public static WriterStyles Create(IWorkbook workbook)
    {
        return new WriterStyles(
            workbook,
            cellStyle: CreateStyle(workbook, isLocked: false, wrapText: true),
            lockedStyle: CreateStyle(workbook, isLocked: true, wrapText: true),
            dateCellStyle: CreateDateStyle(workbook, isLocked: false),
            dateLockedStyle: CreateDateStyle(workbook, isLocked: true));
    }

    /// <summary>按表头颜色返回（缓存）对应的锁定表头样式</summary>
    public ICellStyle HeaderStyleFor(IWorkbook workbook, short color)
    {
        if (_headersByColor.TryGetValue(color, out var style))
        {
            return style;
        }
        var newStyle = CreateStyle(workbook, isLocked: true, wrapText: true);
        var font = workbook.CreateFont();
        font.Color = color;
        newStyle.SetFont(font);
        _headersByColor[color] = newStyle;
        return newStyle;
    }

    /// <summary>根据列锁定状态与值类型返回单元格样式（无附加选项时复用预建样式，保持原性能）</summary>
    public ICellStyle CellStyleFor(object? value, ExcelColumn column)
    {
        if (value is DateTime)
        {
            return column.IsLocked ? _dateLockedStyle : _dateCellStyle;
        }
        return column.IsLocked ? _lockedStyle : _cellStyle;
    }

    /// <summary>
    /// 带附加样式的单元格样式：交替行填充 / 边框 / 数字格式。
    /// 未启用任何附加项时等价于 <see cref="CellStyleFor(object?, ExcelColumn)"/>。
    /// </summary>
    public ICellStyle CellStyleFor(object? value, ExcelColumn column, bool alternateRow, short alternateFillColor, bool border)
    {
        bool hasExtra = alternateRow || border || !string.IsNullOrEmpty(column.NumberFormat);
        if (!hasExtra)
        {
            return CellStyleFor(value, column);
        }

        bool isDate = value is DateTime;
        bool locked = column.IsLocked;
        string key = $"{(locked ? 1 : 0)}|{(isDate ? 1 : 0)}|{(alternateRow ? 1 : 0)}|{alternateFillColor}|{(border ? 1 : 0)}|{column.NumberFormat ?? ""}";
        return _dataStyles.GetOrAdd(key, k => BuildDataStyle(locked, isDate, alternateRow, alternateFillColor, border, column.NumberFormat ?? ""));
    }

    private ICellStyle BuildDataStyle(bool locked, bool isDate, bool alternateRow, short alternateFillColor, bool border, string numberFormat)
    {
        var style = _workbook.CreateCellStyle();
        style.Alignment = HorizontalAlignment.Center;
        style.VerticalAlignment = VerticalAlignment.Center;
        style.IsLocked = locked;
        style.WrapText = true;

        if (isDate)
        {
            style.DataFormat = _workbook.CreateDataFormat().GetFormat("yyyy-mm-dd");
        }
        if (!string.IsNullOrEmpty(numberFormat))
        {
            style.DataFormat = _workbook.CreateDataFormat().GetFormat(numberFormat);
        }
        if (alternateRow && alternateFillColor > 0)
        {
            style.FillForegroundColor = alternateFillColor;
            style.FillPattern = FillPattern.SolidForeground;
        }
        if (border)
        {
            style.BorderTop = BorderStyle.Thin;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
        }
        return style;
    }

    private static ICellStyle CreateStyle(IWorkbook workbook, bool isLocked, bool wrapText)
    {
        var style = workbook.CreateCellStyle();
        style.Alignment = HorizontalAlignment.Center;
        style.VerticalAlignment = VerticalAlignment.Center;
        style.IsLocked = isLocked;
        style.WrapText = wrapText;
        return style;
    }

    private static ICellStyle CreateDateStyle(IWorkbook workbook, bool isLocked)
    {
        var style = workbook.CreateCellStyle();
        style.Alignment = HorizontalAlignment.Center;
        style.VerticalAlignment = VerticalAlignment.Center;
        style.IsLocked = isLocked;
        style.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd");
        return style;
    }
}
