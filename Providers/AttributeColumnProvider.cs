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
            });
        }
        return list.OrderBy(c => c.Order).ThenBy(c => c.Name).ToList();
    }
}
