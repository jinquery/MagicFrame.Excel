using System.Runtime.CompilerServices;
using MagicFrame.Excel.Abstractions;
using NPOI.SS.UserModel;

namespace MagicFrame.Excel.Converters;

/// <summary>
/// 默认单元格值转换器：处理 数值/文本/布尔/日期/空值/公式求值
/// </summary>
public class DefaultCellValueConverter : ICellValueConverter
{
    public static readonly DefaultCellValueConverter Instance = new();

    /// <summary>
    /// 按工作簿缓存公式求值器，避免每个公式单元格都新建求值器导致 O(n²) 性能劣化。
    /// ConditionalWeakTable 不阻止工作簿被 GC，求值器随工作簿一起回收，无泄漏。
    /// </summary>
    private static readonly ConditionalWeakTable<IWorkbook, IFormulaEvaluator> EvaluatorCache = new();

    public object? ReadCell(ICell? cell, Type targetType)
    {
        if (cell == null)
        {
            return ValueConvert.Convert("", targetType);
        }

        object raw;
        switch (cell.CellType)
        {
            case CellType.Numeric:
                if (IsDateCell(cell))
                {
                    raw = cell.DateCellValue!;
                }
                else
                {
                    raw = cell.NumericCellValue;
                }
                break;
            case CellType.String:
                raw = cell.StringCellValue;
                break;
            case CellType.Boolean:
                raw = cell.BooleanCellValue;
                break;
            case CellType.Blank:
                raw = "";
                break;
            case CellType.Error:
                raw = "";
                break;
            case CellType.Formula:
                raw = EvaluateFormula(cell);
                break;
            default:
                raw = "";
                break;
        }
        return ValueConvert.Convert(raw, targetType);
    }

    public string? ToCellText(object? value)
    {
        return value?.ToString();
    }

    public bool IsNumeric(object? value)
    {
        if (value == null)
        {
            return false;
        }
        switch (Type.GetTypeCode(value.GetType()))
        {
            case TypeCode.Byte:
            case TypeCode.SByte:
            case TypeCode.UInt16:
            case TypeCode.UInt32:
            case TypeCode.UInt64:
            case TypeCode.Int16:
            case TypeCode.Int32:
            case TypeCode.Int64:
            case TypeCode.Decimal:
            case TypeCode.Double:
            case TypeCode.Single:
                return true;
            default:
                return false;
        }
    }

    private static bool IsDateCell(ICell cell) => DateUtil.IsCellDateFormatted(cell);

    private static object EvaluateFormula(ICell cell)
    {
        try
        {
            var workbook = cell.Sheet.Workbook;
            // 每个工作簿复用同一个求值器（原子获取，线程安全）
            var evaluator = EvaluatorCache.GetValue(workbook, w => w.GetCreationHelper().CreateFormulaEvaluator()!);
            CellValue cellValue = evaluator.Evaluate(cell);
            if (cellValue == null)
            {
                return "";
            }
            return cellValue.CellType switch
            {
                CellType.Numeric => cellValue.NumberValue,
                CellType.String => cellValue.StringValue,
                CellType.Boolean => cellValue.BooleanValue,
                CellType.Blank => "",
                CellType.Error => cell.CellFormula,
                _ => ""
            };
        }
        catch
        {
            return cell.CellFormula;
        }
    }
}
