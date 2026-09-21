using MagicFrame.Excel.Attributes;
using MagicFrame.Excel.Converters;
using MagicFrame.Excel.Engine;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;
using NPOI.XSSF.UserModel;
using Xunit;

namespace MagicFrame.Excel.Tests;

/// <summary>
/// 值映射测试：实体代码值 ⇄ Excel 显示文本。
/// 支持程序化 <see cref="ExcelColumn.ValueMap"/> 与特性 <c>ValueMappings</c> 两种配置；
/// "unknown" 哨兵键指定异常值兜底代码，未提供时默认 -9999999。
/// </summary>
public class ValueMappingTests
{
    private readonly ExcelEngine _engine = TestHelper.Engine;

    // ---------- 程序化映射 ----------

    [Fact]
    public void Export_Writes_Display_Text_Import_Restores_Code()
    {
        var columns = new[]
        {
            new ExcelColumn { Name = "姓名", Field = "Name", Order = 1 },
            new ExcelColumn
            {
                Name = "性别",
                Field = "Gender",
                Order = 2,
                ValueMap = new Dictionary<object, string> { [0] = "女", [1] = "男" },
            },
        };
        var rows = new[]
        {
            new GenderEntity { Name = "张三", Gender = 1 },
            new GenderEntity { Name = "李四", Gender = 0 },
        };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表", Columns = columns });

        // 导出后单元格显示"男/女"文本
        var sheet = result.Workbook.GetSheetAt(0);
        Assert.Equal("男", sheet.GetRow(1).GetCell(1).StringCellValue);
        Assert.Equal("女", sheet.GetRow(2).GetCell(1).StringCellValue);

        // 导入反查还原为 1/0
        var imported = _engine.Import<GenderEntity>(result.Bytes, new ExcelImportOptions { SheetName = "表", Columns = columns });
        Assert.Equal(2, imported.Count);
        Assert.Equal(1, imported[0].Gender);
        Assert.Equal(0, imported[1].Gender);
    }

    [Fact]
    public void Unmapped_Text_With_Unknown_Key_Returns_Custom_Code()
    {
        var columns = new[]
        {
            new ExcelColumn { Name = "姓名", Field = "Name", Order = 1 },
            new ExcelColumn
            {
                Name = "性别",
                Field = "Gender",
                Order = 2,
                ValueMap = new Dictionary<object, string>
                {
                    [0] = "女",
                    [1] = "男",
                    [ValueMapper.UnknownKey] = "2", // 自定义异常值兜底
                },
            },
        };
        var rows = new[] { new GenderEntity { Name = "张三", Gender = 1 } };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表", Columns = columns });
        var sheet = Assert.IsType<XSSFSheet>(result.Workbook.GetSheetAt(0));
        sheet.GetRow(1).GetCell(1).SetCellValue("未知");
        using var ms = new MemoryStream();
        result.Workbook.Write(ms, true);

        var imported = _engine.Import<GenderEntity>(ms.ToArray(), new ExcelImportOptions { SheetName = "表", Columns = columns });
        Assert.Equal(2, imported[0].Gender); // 自定义 unknown 代码
    }

    [Fact]
    public void Unmapped_Text_Without_Unknown_Key_Returns_Default_Abnormal()
    {
        var columns = new[]
        {
            new ExcelColumn { Name = "姓名", Field = "Name", Order = 1 },
            new ExcelColumn { Name = "性别", Field = "Gender", Order = 2, ValueMap = new Dictionary<object, string> { [0] = "女", [1] = "男" } },
        };
        var rows = new[] { new GenderEntity { Name = "张三", Gender = 1 } };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表", Columns = columns });
        var sheet = Assert.IsType<XSSFSheet>(result.Workbook.GetSheetAt(0));
        sheet.GetRow(1).GetCell(1).SetCellValue("未知");
        using var ms = new MemoryStream();
        result.Workbook.Write(ms, true);

        var imported = _engine.Import<GenderEntity>(ms.ToArray(), new ExcelImportOptions { SheetName = "表", Columns = columns });
        Assert.Equal(ValueMapper.DefaultUnknownValue, imported[0].Gender); // 默认异常值 -9999999
    }

    [Fact]
    public void Blank_Cell_Stays_Default_Not_Abnormal()
    {
        var columns = new[]
        {
            new ExcelColumn { Name = "姓名", Field = "Name", Order = 1 },
            new ExcelColumn { Name = "性别", Field = "Gender", Order = 2, ValueMap = new Dictionary<object, string> { [0] = "女", [1] = "男" } },
        };
        var rows = new[] { new GenderEntity { Name = "张三", Gender = 1 } };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表", Columns = columns });
        var sheet = Assert.IsType<XSSFSheet>(result.Workbook.GetSheetAt(0));
        sheet.GetRow(1).GetCell(1).SetCellValue(""); // 空单元格
        using var ms = new MemoryStream();
        result.Workbook.Write(ms, true);

        var imported = _engine.Import<GenderEntity>(ms.ToArray(), new ExcelImportOptions { SheetName = "表", Columns = columns });
        Assert.Equal(0, imported[0].Gender); // 空单元格 -> 目标默认值 0，而非异常值
    }

    [Fact]
    public void Unmapped_Code_On_Export_Writes_As_Is()
    {
        var columns = new[]
        {
            new ExcelColumn { Name = "姓名", Field = "Name", Order = 1 },
            new ExcelColumn { Name = "性别", Field = "Gender", Order = 2, ValueMap = new Dictionary<object, string> { [0] = "女", [1] = "男" } },
        };
        var rows = new[] { new GenderEntity { Name = "张三", Gender = 9 } }; // 9 不在映射中

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表", Columns = columns });
        var cell = result.Workbook.GetSheetAt(0).GetRow(1).GetCell(1);
        Assert.Equal("9", cell.ToString()); // 按原值写入
    }

    // ---------- 特性映射 ----------

    [Fact]
    public void Attribute_Mapping_RoundTrip_With_Unknown_Key()
    {
        var rows = new List<GenderEntity>
        {
            new() { Name = "张三", Gender = 1 },
            new() { Name = "李四", Gender = 0 },
        };
        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表" });

        var sheet = Assert.IsType<XSSFSheet>(result.Workbook.GetSheetAt(0));
        Assert.Equal("男", sheet.GetRow(1).GetCell(1).StringCellValue);
        Assert.Equal("女", sheet.GetRow(2).GetCell(1).StringCellValue);

        // 改成未映射文本 -> unknown:2 兜底
        sheet.GetRow(1).GetCell(1).SetCellValue("未知性别");
        using var ms = new MemoryStream();
        result.Workbook.Write(ms, true);

        var imported = _engine.Import<GenderEntity>(ms.ToArray(), new ExcelImportOptions { SheetName = "表" });
        Assert.Equal(2, imported[0].Gender);   // unknown 兜底代码
        Assert.Equal(0, imported[1].Gender);   // 正常还原
    }

    [Fact]
    public void Attribute_Without_Unknown_Key_Auto_Adds_Default()
    {
        var rows = new List<AutoUnknownEntity> { new() { Name = "张三", Gender = 1 } };
        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表" });

        var sheet = Assert.IsType<XSSFSheet>(result.Workbook.GetSheetAt(0));
        sheet.GetRow(1).GetCell(1).SetCellValue("未知");
        using var ms = new MemoryStream();
        result.Workbook.Write(ms, true);

        var imported = _engine.Import<AutoUnknownEntity>(ms.ToArray(), new ExcelImportOptions { SheetName = "表" });
        Assert.Equal(ValueMapper.DefaultUnknownValue, imported[0].Gender); // 自动补的 -9999999
    }

    [Fact]
    public void Mapped_Column_Works_With_Locked_And_Group_Header()
    {
        // 分组表头 + 值映射 + 锁定 组合验证
        var columns = new[]
        {
            new ExcelColumn { Name = "姓名", Field = "Name", Order = 1, GroupName = "Base", GroupText = "基础信息" },
            new ExcelColumn
            {
                Name = "性别",
                Field = "Gender",
                Order = 2,
                GroupName = "Base",
                GroupText = "基础信息",
                IsLocked = true,
                ValueMap = new Dictionary<object, string> { [0] = "女", [1] = "男" },
            },
        };
        var rows = new[] { new GenderEntity { Name = "张三", Gender = 0 } };

        using var result = _engine.Export(rows, new ExcelExportOptions { SheetName = "表", HeaderKind = HeaderKinds.Group, Columns = columns });
        var imported = _engine.Import<GenderEntity>(result.Bytes, new ExcelImportOptions { SheetName = "表", HeaderKind = HeaderKinds.Group, Columns = columns });

        Assert.Single(imported);
        Assert.Equal(0, imported[0].Gender);
        // 锁定列样式的 IsLocked 仍生效
        var cell = result.Workbook.GetSheetAt(0).GetRow(2).GetCell(1);
        Assert.True(cell.CellStyle.IsLocked);
    }
}

/// <summary>未提供 unknown 键的特性映射：应自动补 -9999999</summary>
public class AutoUnknownEntity
{
    [ExcelColumn("姓名", Order = 1)]
    public string Name { get; set; } = "";

    [ExcelColumn("性别", Order = 2, ValueMappings = "0:女,1:男")]
    public int Gender { get; set; }
}
