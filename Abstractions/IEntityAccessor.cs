namespace MagicFrame.Excel.Abstractions;

/// <summary>
/// 实体读写访问器：负责按列 Field 读写实体属性值。
/// 默认实现为反射访问器；如需支持索引器实体、字典行等，实现本接口即可。
/// </summary>
public interface IEntityAccessor
{
    /// <summary>
    /// 读取实体某属性值
    /// </summary>
    /// <param name="entity">实体</param>
    /// <param name="field">字段名</param>
    /// <returns></returns>
    object? Read(object entity, string field);

    /// <summary>
    /// 将单元格转换后的值写回实体某属性
    /// </summary>
    /// <param name="entity">实体</param>
    /// <param name="field">实体字段名</param>
    /// <param name="value">单元格内容</param>
    void Write(object entity, string field, object? value);
}
