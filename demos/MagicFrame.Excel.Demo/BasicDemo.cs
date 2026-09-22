using MagicFrame.Excel.Engine;
using MagicFrame.Excel.Model;
using MagicFrame.Excel.Options;

namespace MagicFrame.Excel.Demo;

/// <summary>
/// 基础功能演示：单表头、公式、整列锁定、下拉、隐藏列、分组表头、多 Sheet、Base64、0 值空白
/// </summary>
public static class BasicDemo
{
    public static void Run(string outputDir)
    {
        var engine = ExcelEngine.CreateDefault();
        Directory.CreateDirectory(outputDir);

        Console.WriteLine("==================== 基础功能演示 ====================");
        Console.WriteLine();

        // 1. 单表头 + 公式 + 整列锁定 + 下拉 + 隐藏列
        Console.WriteLine("[1] 单表头导出（公式/锁定/下拉/隐藏列）→ employee.xlsx");
        var employees = new List<Employee>
        {
            new() { Name = "张三", Department = "研发", BaseSalary = 10000m, Factor = 1.2m, HireDate = new DateTime(2020, 3, 15), Remark = "组长" },
            new() { Name = "李四", Department = "测试", BaseSalary = 8000m, Factor = 1.0m, HireDate = new DateTime(2019, 7, 1), Remark = "内推" },
            new() { Name = "王五", Department = "产品", BaseSalary = 12000m, Factor = 0.9m, HireDate = new DateTime(2022, 11, 20), Remark = null },
        };
        string employeeFile = Path.Combine(outputDir, "employee.xlsx");
        engine.ExportToFile(employees, employeeFile, new ExcelExportOptions
        {
            SheetName = "员工表",
            Protection = new SheetProtectionOptions { Password = "1234", AllowInsertRows = true, AllowDeleteRows = true },
        });
        Console.WriteLine($"    已生成: {employeeFile}");

        // 2. 解析回来并打印
        Console.WriteLine("[2] 解析 employee.xlsx → 控制台");
        var back = engine.ImportFile<Employee>(employeeFile, new ExcelImportOptions { SheetName = "员工表" });
        PrintGrid(back.Select(e => new object?[]
        {
            e.Name, e.Department, e.BaseSalary, e.Factor, e.Total, e.HireDate.ToString("yyyy-MM-dd"), e.Remark
        }).ToList(), new[] { "姓名", "部门", "基础工资", "系数", "应发合计(公式)", "入职日期", "备注" });
        Console.WriteLine();

        // 3. 分组表头
        Console.WriteLine("[3] 分组表头导出 → assessment.xlsx");
        var assessments = new List<Assessment>
        {
            new() { Name = "张三", Department = "研发", Performance = 90, Attendance = 95, Grade = "A" },
            new() { Name = "李四", Department = "测试", Performance = 70, Attendance = 80, Grade = "B" },
            new() { Name = "王五", Department = "产品", Performance = 85, Attendance = 90, Grade = "A" },
        };
        string assessmentFile = Path.Combine(outputDir, "assessment.xlsx");
        engine.ExportToFile(assessments, assessmentFile, new ExcelExportOptions
        {
            SheetName = "考核表",
            HeaderKind = HeaderKinds.Group,
            EnableBorders = true, // 整表（表头+数据）加细边框
        });
        var assessmentBack = engine.ImportFile<Assessment>(assessmentFile, new ExcelImportOptions { SheetName = "考核表", HeaderKind = HeaderKinds.Group });
        PrintGrid(assessmentBack.Select(a => new object?[] { a.Name, a.Department, a.Performance, a.Attendance, a.Total, a.Grade }).ToList(),
            new[] { "姓名", "部门", "业绩", "出勤", "总分(公式)", "评级" });
        Console.WriteLine();

        // 4. 多 Sheet
        Console.WriteLine("[4] 多 Sheet 导出 → multi_sheet.xlsx");
        var specs = new[]
        {
            SheetSpec.Of("员工表", employees),
            SheetSpec.Of("考核表", assessments),
        };
        string multiFile = Path.Combine(outputDir, "multi_sheet.xlsx");
        File.WriteAllBytes(multiFile, engine.ExportSheets(specs).Bytes);
        Console.WriteLine($"    已生成: {multiFile}（{2} 个 Sheet）");
        Console.WriteLine();

        // 5. 0 值空白
        Console.WriteLine("[5] 0 值空白导出 → zero_blank.xlsx");
        var products = new List<Product>
        {
            new() { Name = "苹果", Sales = 0 },
            new() { Name = "香蕉", Sales = 120 },
        };
        string productFile = Path.Combine(outputDir, "zero_blank.xlsx");
        engine.ExportToFile(products, productFile);
        Console.WriteLine($"    已生成: {productFile}（销量 0 的单元格留空）");
        Console.WriteLine();

        // 6. Base64（供前端下载）
        Console.WriteLine("[6] Base64 导出（前 60 字符）");
        string base64 = engine.ExportBase64(employees);
        Console.WriteLine($"    {base64[..Math.Min(60, base64.Length)]}...");
        var fromBase64 = engine.Import<Employee>(base64);
        Console.WriteLine($"    从 Base64 解析回 {fromBase64.Count} 行");
        Console.WriteLine();

        // 7. 导入列对应校验
        Console.WriteLine("[7] 导入列对应校验");
        using (var exported = engine.Export(employees, new ExcelExportOptions { SheetName = "员工表" }))
        {
            // 人为追加一列无法对应实体列 + 清空"备注"表头（隐藏列，缺失应被忽略）
            var sheet = (NPOI.XSSF.UserModel.XSSFSheet)exported.Workbook.GetSheetAt(0);
            sheet.GetRow(0).CreateCell(7).SetCellValue("神秘列");
            sheet.GetRow(1).CreateCell(7).SetCellValue("x");
            using var ms = new MemoryStream();
            exported.Workbook.Write(ms, true);

            var importOptions = new ExcelImportOptions { SheetName = "员工表" };
            var validated = engine.Import<Employee>(ms.ToArray(), importOptions);
            Console.WriteLine($"    解析 {validated.Count} 行，校验信息：");
            foreach (var msg in importOptions.Errors)
            {
                Console.WriteLine($"      - {msg}");
            }
        }
        Console.WriteLine();

        // 8. 值映射（代码 <-> 显示文本）
        Console.WriteLine("[8] 值映射（0/1/2 <-> 女/男/其他）→ value_map.xlsx");
        var genderRows = new List<GenderRow>
        {
            new() { Name = "张三", Gender = 1 },
            new() { Name = "李四", Gender = 0 },
        };
        var genderColumns = new[]
        {
            new MagicFrame.Excel.Model.ExcelColumn { Name = "姓名", Field = "Name", Order = 1 },
            new MagicFrame.Excel.Model.ExcelColumn
            {
                Name = "性别", Field = "Gender", Order = 2,
                // unknown 哨兵键指定"异常值"兜底代码（未映射文本 -> 2）；缺省为 -9999999
                ValueMap = new Dictionary<object, string>
                {
                    [0] = "女",
                    [1] = "男",
                    [MagicFrame.Excel.Converters.ValueMapper.UnknownKey] = "2",
                },
            },
        };
        string valueMapFile = Path.Combine(outputDir, "value_map.xlsx");
        engine.ExportToFile(genderRows, valueMapFile, new ExcelExportOptions { SheetName = "映射表", Columns = genderColumns });
        // 人为把一格改成未映射文本，验证导入兜底为特殊值 2
        using (var exported = engine.Export(genderRows, new ExcelExportOptions { SheetName = "映射表", Columns = genderColumns }))
        {
            ((NPOI.XSSF.UserModel.XSSFSheet)exported.Workbook.GetSheetAt(0)).GetRow(2).GetCell(1).SetCellValue("未知");
            using var ms = new MemoryStream();
            exported.Workbook.Write(ms, true);
            var mappedBack = engine.Import<GenderRow>(ms.ToArray(), new ExcelImportOptions { SheetName = "映射表", Columns = genderColumns });
            Console.WriteLine($"    已生成: {valueMapFile}；解析回 {mappedBack.Count} 行，Gender = [{string.Join(", ", mappedBack.Select(g => g.Gender))}]（张三=1男，李四=2其他[未映射兜底]）");
        }
        Console.WriteLine();

        Console.WriteLine("==================== 基础演示结束 ====================");
        Console.WriteLine();
    }

    private static void PrintGrid(List<object?[]> rows, string[] headers)
    {
        var widths = headers.Select((h, i) => Math.Max(h.Length, rows.Count > 0 ? rows.Max(r => r[i]?.ToString()?.Length ?? 0) : 0)).ToArray();

        void PrintLine()
        {
            Console.WriteLine("    " + string.Join("-+-", widths.Select(w => new string('-', Math.Max(w, 1)))));
        }

        PrintLine();
        Console.WriteLine("    " + string.Join(" | ", headers.Select((h, i) => h.PadRight(widths[i]))));
        PrintLine();
        foreach (var row in rows)
        {
            Console.WriteLine("    " + string.Join(" | ", row.Select((v, i) => (v?.ToString() ?? "").PadRight(widths[i]))));
        }
        PrintLine();
    }
}
