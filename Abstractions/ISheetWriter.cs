using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using NPOI.SS.UserModel;

namespace MagicFrame.Excel.Abstractions;

/// <summary>
/// Sheet 写入策略：把一个实体集合写入一个工作表。
/// 内置实现：单表头 <see cref="Writers.SingleHeaderSheetWriter"/>、分组表头 <see cref="Writers.GroupHeaderSheetWriter"/>。
/// 如需新的表头布局（如固定行、表头样式定制等），实现本接口并通过 Builder 注册到新的 HeaderKind 即可。
/// </summary>
public interface ISheetWriter
{
    /// <summary>表头占用行数（用于冻结窗格与数据起始行计算）</summary>
    int HeaderRowCount { get; }

    /// <summary>
    /// 写入工作表
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="workbook">workbook对象</param>
    /// <param name="columns">列</param>
    /// <param name="accessor">实体读写访问器：负责按列 Field 读写实体属性值</param>
    /// <param name="converter">实体字段到单元格内容数据转换器</param>
    /// <param name="formulaResolver">解析公式模板</param>
    /// <param name="rows">待写入的数据集</param>
    /// <param name="options">Excel的配置</param>
    void Write(
        ISheet sheet,
        IWorkbook workbook,
        IReadOnlyList<ExcelColumn> columns,
        IEntityAccessor accessor,
        ICellValueConverter converter,
        IFormulaResolver formulaResolver,
        IEnumerable<object> rows,
        ExcelExportOptions options);
}
