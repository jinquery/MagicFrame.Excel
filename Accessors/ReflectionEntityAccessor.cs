using System.Collections.Concurrent;
using System.Linq.Expressions;
using System.Reflection;
using MagicFrame.Excel.Abstractions;
using MagicFrame.Excel.Converters;

namespace MagicFrame.Excel.Accessors;

/// <summary>
/// 反射实体访问器：按属性名读写 POCO 属性，写值时自动做类型转换。
/// <para>
/// 性能优化：
/// - 读写属性编译为表达式树委托并静态缓存（按 (Type, Field) 只编译一次）；
/// - <see cref="GetWriter(Type, string)"/> 返回"转换+赋值"合并委托，供读取器按列预解析，避免每格字典查找；
/// - <see cref="GetFactory(Type)"/> 缓存行对象工厂，避免每行 Activator.CreateInstance；
/// - 表达式编译产物为托管委托（DynamicMethod），无动态程序集、GC 可回收，静态缓存不随调用次数增长，无泄漏。
/// </para>
/// </summary>
public class ReflectionEntityAccessor : IEntityAccessor
{
    public static readonly ReflectionEntityAccessor Instance = new();

    // 已编译 getter/writer/setter 委托缓存（null 表示该字段不可读写）
    private static readonly ConcurrentDictionary<(Type Type, string Field), Func<object, object?>?> GetterCache = new();
    private static readonly ConcurrentDictionary<(Type Type, string Field), Action<object, object?>?> WriterCache = new();
    private static readonly ConcurrentDictionary<(Type Type, string Field), Action<object, object?>?> SetterCache = new();

    // 属性类型缓存（避免在行列循环中重复反射）
    private static readonly ConcurrentDictionary<(Type Type, string Field), Type> TypeCache = new();

    // 行对象工厂缓存
    private static readonly ConcurrentDictionary<Type, Func<object>> FactoryCache = new();

    public object? Read(object entity, string field)
    {
        var getter = GetterCache.GetOrAdd((entity.GetType(), field), k => BuildGetter(k.Type, k.Field));
        return getter?.Invoke(entity);
    }

    public void Write(object entity, string field, object? value)
    {
        var writer = WriterCache.GetOrAdd((entity.GetType(), field), k => BuildWriter(k.Type, k.Field));
        writer?.Invoke(entity, value);
    }

    /// <summary>
    /// 解析字段对应的属性类型（缓存，避免重复反射）；属性不存在时返回 string
    /// </summary>
    public static Type GetPropertyType(Type entityType, string field)
        => TypeCache.GetOrAdd((entityType, field), k => ResolveType(k.Type, k.Field));

    /// <summary>
    /// 获取"转换+赋值"委托：entity, rawValue -> 写属性（内部经 <see cref="ValueConvert.Convert"/> 转目标类型）。
    /// 供读取器在循环前按列预解析，调用时直接 Invoke，省去每格字典查找（A5/A6）。
    /// </summary>
    public static Action<object, object?>? GetWriter(Type entityType, string field)
        => WriterCache.GetOrAdd((entityType, field), k => BuildWriter(k.Type, k.Field));

    /// <summary>
    /// 获取纯赋值委托：entity, value -> prop = (T)value（不做类型转换，值须已为目标类型）。
    /// 配合默认转换器使用时，可省去每格一次 ValueConvert（消除二次转换）。
    /// </summary>
    public static Action<object, object?>? GetSetter(Type entityType, string field)
        => SetterCache.GetOrAdd((entityType, field), k => BuildSetter(k.Type, k.Field));

    /// <summary>
    /// 获取编译的 getter 委托（供导出路径按列预解析，避免每格 GetOrAdd）
    /// </summary>
    public static Func<object, object?>? GetGetter(Type entityType, string field)
        => GetterCache.GetOrAdd((entityType, field), k => BuildGetter(k.Type, k.Field));

    /// <summary>
    /// 获取行对象工厂（缓存编译的 new），避免每行 Activator.CreateInstance（A4）
    /// </summary>
    public static Func<object> GetFactory(Type entityType)
        => FactoryCache.GetOrAdd(entityType, t =>
            Expression.Lambda<Func<object>>(Expression.Convert(Expression.New(t), typeof(object))).Compile());

    private static Type ResolveType(Type type, string field)
        => type.GetProperty(field, BindingFlags.Public | BindingFlags.Instance)?.PropertyType ?? typeof(string);

    /// <summary>
    /// 编译 getter 委托：entity -> (object?)prop
    /// </summary>
    private static Func<object, object?>? BuildGetter(Type type, string field)
    {
        var prop = type.GetProperty(field, BindingFlags.Public | BindingFlags.Instance);
        if (prop == null || !prop.CanRead)
        {
            return null;
        }
        var entity = Expression.Parameter(typeof(object), "entity");
        var body = Expression.Convert(
            Expression.Property(Expression.Convert(entity, type), prop),
            typeof(object));
        return Expression.Lambda<Func<object, object?>>(body, entity).Compile();
    }

    /// <summary>
    /// 编译"转换+赋值"委托：entity, rawValue -> prop = (T)ValueConvert.Convert(rawValue, T)
    /// </summary>
    private static Action<object, object?>? BuildWriter(Type type, string field)
    {
        var prop = type.GetProperty(field, BindingFlags.Public | BindingFlags.Instance);
        if (prop == null || !prop.CanWrite)
        {
            return null;
        }
        var propertyType = prop.PropertyType;
        var entity = Expression.Parameter(typeof(object), "entity");
        var value = Expression.Parameter(typeof(object), "value");
        var convert = Expression.Call(
            typeof(ValueConvert).GetMethod(nameof(ValueConvert.Convert), new[] { typeof(object), typeof(Type) })!,
            value,
            Expression.Constant(propertyType));
        var assign = Expression.Assign(
            Expression.Property(Expression.Convert(entity, type), prop),
            Expression.Convert(convert, propertyType));
        return Expression.Lambda<Action<object, object?>>(assign, entity, value).Compile();
    }

    /// <summary>
    /// 编译纯赋值委托：entity, value -> prop = (T)value（值须已为目标类型，不做转换）
    /// </summary>
    private static Action<object, object?>? BuildSetter(Type type, string field)
    {
        var prop = type.GetProperty(field, BindingFlags.Public | BindingFlags.Instance);
        if (prop == null || !prop.CanWrite)
        {
            return null;
        }
        var entity = Expression.Parameter(typeof(object), "entity");
        var value = Expression.Parameter(typeof(object), "value");
        var assign = Expression.Assign(
            Expression.Property(Expression.Convert(entity, type), prop),
            Expression.Convert(value, prop.PropertyType));
        return Expression.Lambda<Action<object, object?>>(assign, entity, value).Compile();
    }
}
