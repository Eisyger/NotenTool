using System.Diagnostics;
using System.Text.Json;
using Application.DTO;
using Domain.Entity;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Application.Services;

public class ExcelMergeExportService
{
    public static byte[] GetFile(List<ExamResultDto> examResults)
    {
        using FileStream fs = new("test.json", FileMode.Create);
        fs.Write(JsonSerializer.SerializeToUtf8Bytes(examResults));
        Console.WriteLine(JsonSerializer.Serialize(examResults));
        
        Environment.SetEnvironmentVariable("NPOI_FONT_PATH", "");

        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Ergebnisse");

        // zentrale Struktur vorbereiten
        var structure = BuildStructure(examResults);

        SetupHeaders(sheet, structure);
        FillSheet(sheet, examResults, structure);

        using var stream = new MemoryStream();
        workbook.Write(stream, true);
        return stream.ToArray();
    }
    
    private static List<(string ExamName, List<string> Examiners)> BuildStructure(List<ExamResultDto> examResults)
    {
        return examResults
            .SelectMany(x => x.Exam)
            .GroupBy(e => e.ExamName)
            .OrderBy(g => g.Key)
            .Select(g => (
                ExamName: g.Key,
                Examiners: g
                    .Select(e => e.Examiner)
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList()
            ))
            .ToList();
    }
    private static void SetupHeaders(
        ISheet sheet,
        List<(string ExamName, List<string> Examiners)> structure)
    {
        var headerStyle = CreateHeaderStyle(sheet.Workbook);
        
        var row = sheet.CreateRow(0);
        var col = 0;

        var firstCol = row.CreateCell(col++);
        firstCol.CellStyle = headerStyle;
        firstCol.SetCellValue("Vorname");
        sheet.SetColumnWidth(col-1, ("Vorname".Length + 2) * 256);

        var secondCol = row.CreateCell(col++);
        secondCol.CellStyle = headerStyle;
        secondCol.SetCellValue("Nachname");
        sheet.SetColumnWidth(col-1, ("Nachname".Length + 2) * 256);

        foreach (var exam in structure)
        {
            var examCol = row.CreateCell(col++);
            examCol.CellStyle = headerStyle;
            examCol.SetCellValue(exam.ExamName);
            
            sheet.SetColumnWidth(col-1, (exam.ExamName.Length + 2) * 256);

            foreach (var examiner in exam.Examiners)
            {
                var examinerCol = row.CreateCell(col++);
                examinerCol.CellStyle = headerStyle;
                examinerCol.SetCellValue(examiner);
                
                sheet.SetColumnWidth(col-1, (examiner.Length + 2) * 256);
            }
        }

        sheet.CreateFreezePane(0, 1);
    }

    private static void FillSheet(
        ISheet sheet,
        List<ExamResultDto> examResults,
        List<(string ExamName, List<string> Examiners)> structure)
    {
        var numericStyle = sheet.Workbook.CreateCellStyle();
        
        numericStyle.Alignment = HorizontalAlignment.Left;
        numericStyle.VerticalAlignment = VerticalAlignment.Center;

        var dataFormat = sheet.Workbook.CreateDataFormat();
        numericStyle.DataFormat = dataFormat.GetFormat("0.0");  // z. B. "1,7" / "2,3"
        
        var rowIndex = 1;

        foreach (var student in examResults)
        {
            var row = sheet.CreateRow(rowIndex++);
            var col = 0;

            row.CreateCell(col++).SetCellValue(student.FirstName);
            row.CreateCell(col++).SetCellValue(student.LastName);

            // 🔥 einmal Lookup bauen → O(1)
            var lookup = student.Exam
                .GroupBy(e => (e.ExamName, e.Examiner))
                .ToDictionary(g => g.Key, g => g.First());

            foreach (var exam in structure)
            {
                var hasExam = student.Exam.Any(e => e.ExamName == exam.ExamName);

                row.CreateCell(col++).SetCellValue(hasExam ? exam.ExamName : "");

                foreach (var examiner in exam.Examiners)
                {
                    var cell = row.CreateCell(col++);

                    if (lookup.TryGetValue((exam.ExamName, examiner), out var e))
                    {
                        if (e.Grade > 0.0f)
                        {
                            cell.SetCellValue(e.Grade);
                            cell.CellStyle = numericStyle;
                        }
                    }
                }
            }
        }
    }
    
    private static ICellStyle CreateHeaderStyle(IWorkbook workbook)
    {
        var style = workbook.CreateCellStyle();
        
        // Background color
        style.FillForegroundColor = IndexedColors.CornflowerBlue.Index;
        style.FillPattern = FillPattern.SolidForeground;
        
        // Border
        style.BorderTop = BorderStyle.Thin;
        style.BorderBottom = BorderStyle.Thin;
        style.BorderLeft = BorderStyle.Thin;
        style.BorderRight = BorderStyle.Thin;
        
        // Alignment
        style.Alignment = HorizontalAlignment.Center;
        style.VerticalAlignment = VerticalAlignment.Center;
        
        return style;
    }
}
