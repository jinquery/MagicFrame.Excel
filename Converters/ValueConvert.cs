using System.Globalization;
using System.Text.RegularExpressions;

namespace MagicFrame.Excel.Converters;

/// <summary>
/// 值类型转换工具：把 Excel 单元格读出的原始值转换为目标属性类型
/// </summary>
public static class ValueConvert
{
    /// <summary>
    /// 将值转换为目标类型（支持 数值/布尔/日期/可空/枚举/字符串）
    /// </summary>
    public static object? Convert(object? value, Type targetType)
    {
        if (value == null || value is DBNull)
        {
            return targetType.IsValueType ? Activator.CreateInstance(targetType) : null;
        }

        var underlying = Nullable.GetUnderlyingType(targetType);
        Type effective = underlying ?? targetType;

        if (effective == typeof(string))
        {
            return value.ToString();
        }

        if (effective.IsEnum)
        {
            if (value is string s && int.TryParse(s, out _))
            {
                return Enum.ToObject(effective, int.Parse(s));
            }
            return Enum.Parse(effective, value.ToString()!, ignoreCase: true);
        }

        if (effective == typeof(bool))
        {
            return ToBool(value);
        }

        if (effective == typeof(DateTime))
        {
            return ToDateTime(value);
        }

        if (IsNumericType(effective))
        {
            return ToNumeric(value, effective);
        }

        // 其它类型（Guid 等）走 ChangeType
        try
        {
            return System.Convert.ChangeType(value, effective, CultureInfo.InvariantCulture);
        }
        catch
        {
            return value;
        }
    }

    private static bool ToBool(object value)
    {
        string text = value.ToString()?.Trim().ToUpperInvariant() ?? "";
        if ("Y".Equals(text) || "1".Equals(text) || "TRUE".Equals(text) || "是".Equals(text))
        {
            return true;
        }
        if (value is bool b)
        {
            return b;
        }
        if (value is double d)
        {
            return Math.Abs(d) > 0;
        }
        return false;
    }

    private static DateTime ToDateTime(object value)
    {
        if (value is DateTime dt)
        {
            return dt;
        }
        if (value is double d2)
        {
            return DateTime.FromOADate(d2);
        }
        return DateTime.TryParse(value.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
            ? parsed
            : DateTime.MinValue;
    }

    private static object ToNumeric(object value, Type type)
    {
        if (value is double d)
        {
            return System.Convert.ChangeType(d, type, CultureInfo.InvariantCulture);
        }
        if (value is int || value is long || value is decimal || value is float)
        {
            return System.Convert.ChangeType(value, type, CultureInfo.InvariantCulture);
        }
        string text = value.ToString() ?? "";
        // 去除除数字、小数点、负号外的字符（与旧实现一致）
        string cleaned = Regex.Replace(text, @"[^\d.\-]", "");
        if (string.IsNullOrWhiteSpace(cleaned))
        {
            return Activator.CreateInstance(type)!;
        }
        try
        {
            double parsed = double.Parse(cleaned, CultureInfo.InvariantCulture);
            return System.Convert.ChangeType(parsed, type, CultureInfo.InvariantCulture);
        }
        catch
        {
            return Activator.CreateInstance(type)!;
        }
    }

    private static bool IsNumericType(Type type)
    {
        switch (Type.GetTypeCode(type))
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
}
