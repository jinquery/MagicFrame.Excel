using System.Collections.Concurrent;
using System.Reflection;
using MagicFrame.Excel.Abstractions;
using MagicFrame.Excel.Attributes;
using MagicFrame.Excel.Model;

namespace MagicFrame.Excel.Providers;

/// <summary>
/// 特性列提供器：通过反射读取实体上的 <see cref="ExcelColumnAttribute"/> 生成列定义。
/// 列定义按实体类型缓存（ConcurrentDictionary&lt;Type, IReadOnlyList&lt;ExcelColumn&gt;&gt;），
/// 返回时逐列 <see cref="ExcelColumn.Clone"/>，避免外部修改污染缓存（D1）。
/// </summary>
public class AttributeColumnProvider : IColumnProvider
{
    public static readonly AttributeColumnProvider Instance = new();

    private static readonly ConcurrentDictionary<Type, IReadOnlyList<ExcelColumn>> Cache = new();

    public IReadOnlyList<ExcelColumn> GetColumns(Type entityType)
    {
        var cached = Cache.GetOrAdd(entityType, Build);
        // 返回克隆，避免调用方修改列定义污染缓存
        return cached.Select(c => c.Clone()).ToList();
    }

    private static IReadOnlyList<ExcelColumn> Build(Type entityType)
    {
        var list = new List<ExcelColumn>();
        foreach (var prop in entityType.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            var att = prop.GetCustomAttribute<ExcelColumnAttribute>();
            if (att == null)
            {
                continue;
            }
            list.Add(new ExcelColumn
            {
                Name = att.Name,
                Field = string.IsNullOrWhiteSpace(att.Field) ? prop.Name : att.Field!,
                Order = att.Order,
                Visible = att.Visible,
                IsLocked = att.IsLocked,
                Formula = att.Formula ?? "",
                ZeroShowWhiteSpace = att.ZeroShowWhiteSpace,
                HeaderColor = att.HeaderColor,
                GroupName = att.GroupName ?? "",
                GroupText = att.GroupText ?? "",
                DropdownOptions = att.DropdownOptions,
                Width = att.Width > 0 ? att.Width : null,
                Alignment = att.Alignment,
                WriteAsNumeric = att.WriteAsNumeric,
                NumberFormat = att.NumberFormat ?? "",
                ValidationKind = att.ValidationKind,
                ValidationFormula1 = att.ValidationFormula1 ?? "",
                ValidationFormula2 = att.ValidationFormula2 ?? "",
                ValueMap = ParseValueMappings(att.ValueMappings),
            });
        }
        return list.OrderBy(c => c.Order).ThenBy(c => c.Name).ToList();
    }

    /// <summary>
    /// 解析 "0:女,1:男,unknown:2" -> {0="女",1="男", unknown="2"}。
    /// 纯数字键按 long，其余按 string；"unknown" 为异常值哨兵键。
    /// 未提供 unknown 时自动补 "unknown:-9999999"（异常值默认）。
    /// </summary>
    private static Dictionary<object, string>? ParseValueMappings(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }
        var map = new Dictionary<object, string>();
        bool hasUnknown = false;
        foreach (var segment in text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            int colon = segment.IndexOf(':');
            if (colon <= 0)
            {
                continue;
            }
            string key = segment[..colon].Trim();
            string value = segment[(colon + 1)..].Trim();
            if (string.Equals(key, Converters.ValueMapper.UnknownKey, StringComparison.OrdinalIgnoreCase))
            {
                map[Converters.ValueMapper.UnknownKey] = value;
                hasUnknown = true;
                continue;
            }
            object keyObj = long.TryParse(key, out long num) ? num : key;
            map[keyObj] = value;
        }
        if (!hasUnknown)
        {
            map[Converters.ValueMapper.UnknownKey] = Converters.ValueMapper.DefaultUnknownValue.ToString();
        }
        return map.Count == 0 ? null : map;
    }
}
