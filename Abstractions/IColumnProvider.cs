using MagicFrame.Excel.Model;

namespace MagicFrame.Excel.Abstractions;

/// <summary>
/// 列元数据提供器：为指定实体类型提供列定义。
/// 默认实现为基于 <see cref="Attributes.ExcelColumnAttribute"/> 的特性发现；
/// 如需扩展（配置文件、数据库、前端下发等），实现本接口并在 Builder 中注册即可，核心代码无需改动。
/// </summary>
public interface IColumnProvider
{
    /// <summary>
    /// 获取实体类型的列定义（应按 <see cref="ExcelColumn.Order"/> 排序）
    /// </summary>
    /// <param name="entityType">实体类型</param>
    /// <returns></returns>
    IReadOnlyList<ExcelColumn> GetColumns(Type entityType);
}
