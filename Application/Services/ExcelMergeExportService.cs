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
        var row = sheet.CreateRow(0);
        var col = 0;

        row.CreateCell(col++).SetCellValue("Vorname");
        row.CreateCell(col++).SetCellValue("Nachname");

        foreach (var exam in structure)
        {
            row.CreateCell(col++).SetCellValue(exam.ExamName);

            foreach (var examiner in exam.Examiners)
            {
                row.CreateCell(col++).SetCellValue(examiner);
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
                        cell.SetCellValue(e.Grade);
                        cell.CellStyle = numericStyle;
                    }
                }
            }
        }
    }
}
