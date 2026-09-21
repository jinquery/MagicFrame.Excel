using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;

namespace MagicFrame.Excel.Abstractions;

/// <summary>
/// Sheet 读取策略：把一个工作表解析为实体集合。
/// 内置实现：单表头 <see cref="Readers.SingleHeaderSheetReader"/>、分组表头 <see cref="Readers.GroupHeaderSheetReader"/>。
/// 如需新的表头布局，实现本接口并注册到对应 HeaderKind。
/// </summary>
public interface ISheetReader
{
    /// <summary>
    /// 读取工作表
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="columns">列定义</param>
    /// <param name="entityType">实体类型 T</param>
    /// <param name="accessor">实体访问器</param>
    /// <param name="converter">单元格值转换器</param>
    /// <param name="options">导入选项</param>
    /// <returns>实体集合（object 装箱，由引擎转型为 T）</returns>
    IReadOnlyList<object> Read(
        ISheet sheet,
        IReadOnlyList<ExcelColumn> columns,
        Type entityType,
        IEntityAccessor accessor,
        ICellValueConverter converter,
        ExcelImportOptions options);
}
