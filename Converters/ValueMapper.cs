using MagicFrame.Excel.Model;

namespace MagicFrame.Excel.Converters;

/// <summary>
/// 值映射工具：实体代码值 ⇄ Excel 显示文本。
/// <para>
/// 典型场景：字段存 0/1，Excel 显示 女/男；导入时 男 → 1、女 → 0。
/// </para>
/// <list type="bullet">
/// <item>导出：<see cref="TryGetDisplayText"/> 把实体值换成显示文本写单元格。</item>
/// <item>导入：<see cref="ResolveValue"/> 把显示文本反查回实体值；未命中的非空文本回退到
/// "异常值"——优先取 <see cref="UnknownKey"/>（"unknown"）哨兵键指定的代码，未提供则用
/// <see cref="DefaultUnknownValue"/>（-9999999）。空单元格保持原样（转成目标默认值）。</item>
/// </list>
/// </summary>
public static class ValueMapper
{
    /// <summary>ValueMap 中"异常值"兜底代码的哨兵键</summary>
    public const string UnknownKey = "unknown";

    /// <summary>未提供 unknown 哨兵键时使用的默认异常值</summary>
    public const long DefaultUnknownValue = -9999999L;

    /// <summary>
    /// 实体值 -> 显示文本（导出）。未命中返回 false。
    /// 数值键自动做跨类型归一化（int/byte/short/long 等按 Int64 比较）；"unknown"哨兵键不参与导出映射。
    /// </summary>
    public static bool TryGetDisplayText(ExcelColumn column, object? value, out string? displayText)
    {
        displayText = null;
        var map = column.ValueMap;
        if (map == null || value == null)
        {
            return false;
        }
        if (value is string s && string.Equals(s, UnknownKey, StringComparison.OrdinalIgnoreCase))
        {
            return false; // 哨兵键不导出
        }
        if (map.TryGetValue(value, out displayText))
        {
            return true;
        }
        if (IsNumeric(value))
        {
            long n = ToInt64(value);
            foreach (var kv in map)
            {
                if (IsNumeric(kv.Key) && ToInt64(kv.Key) == n)
                {
                    displayText = kv.Value;
                    return true;
                }
            }
        }
        return false;
    }

    /// <summary>
    /// 显示文本 -> 实体值（导入）。命中返回对应实体值（转为目标类型）；
    /// 未命中的非空文本回退到"异常值"（unknown 哨兵键代码，缺省 -9999999）；空单元格原样返回。
    /// </summary>
    public static object? ResolveValue(ExcelColumn column, object? cellValue, Type targetType)
    {
        var map = column.ValueMap;
        if (map == null)
        {
            return cellValue;
        }

        string? text = cellValue?.ToString();
        // 1) 按显示文本反查（男 -> 1、女 -> 0）
        foreach (var kv in map)
        {
            if (kv.Key is string key && string.Equals(key, UnknownKey, StringComparison.OrdinalIgnoreCase))
            {
                continue; // 哨兵键不参与反查
            }
            if (kv.Value.Equals(cellValue) || (text != null && kv.Value == text))
            {
                return ValueConvert.Convert(kv.Key, targetType);
            }
        }
        // 2) 单元格本身就是代码值（如数值 0/1 或文本 "0"）-> 直接返回
        foreach (var kv in map)
        {
            if (kv.Key is string key && string.Equals(key, UnknownKey, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }
            if (kv.Key.Equals(cellValue) || (text != null && kv.Key.ToString() == text))
            {
                return ValueConvert.Convert(kv.Key, targetType);
            }
        }

        // 3) 未映射的非空文本 -> "异常值"兜底（unknown 哨兵键，缺省 -9999999）；空单元格转目标默认值
        if (!string.IsNullOrEmpty(text))
        {
            object fallback = map.TryGetValue(UnknownKey, out string? code) ? code : DefaultUnknownValue;
            return ValueConvert.Convert(fallback, targetType);
        }
        return ValueConvert.Convert(cellValue, targetType);
    }

    private static bool IsNumeric(object value)
        => value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal;

    private static long ToInt64(object value)
        => Convert.ToInt64(value);
}
