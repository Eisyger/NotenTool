using System.Globalization;
using Application.DTO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Domain.Entity;

namespace Application.Services;

public class ExcelParseService
{
    public async Task<Course> ParseTemplate(Stream fileStream)
    {
        return await Task.Run(() =>
        {
            // Font Suche deaktivieren für WASM
            Environment.SetEnvironmentVariable("NPOI_FONT_PATH", "");
            
            var workbook = new XSSFWorkbook(fileStream);
            
            var data = new Course
            {
                Students = ParseStudentSheet(workbook.GetSheet("Teilnehmer")),
                Exams = ParseExamSheet(workbook.GetSheet("Prüfungen"))
            };
            
            return data;
        });
    }

    public static async Task<List<ExamResultDto>> ParseGrades(Dictionary<string, byte[]> excelTables)
    {
        return await Task.Run(() =>
        {
            Environment.SetEnvironmentVariable("NPOI_FONT_PATH", "");

            var examResults = new List<ExamResultDto>();

            foreach (var (examiner, bytes) in excelTables)
            {
                using var stream = new MemoryStream(bytes);
                var workbook = new XSSFWorkbook(stream);

                for (var i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);
                    ParseExamResults(sheet, examResults, examiner);
                }
            }

            return examResults;
        });
    }

    private static void ParseExamResults(ISheet? sheet, List<ExamResultDto> examResults, string examiner)
    {
        if (sheet is null) return;

        // schneller Zugriff
        var lookup = examResults.ToDictionary(
            x => $"{x.FirstName}|{x.LastName}|{x.Club}",
            x => x
        );

        for (var rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null || IsRowEmpty(row)) continue;

            var firstName = GetCellValue(row.GetCell(0));
            var lastName  = GetCellValue(row.GetCell(1));
            var club      = GetCellValue(row.GetCell(7));

            float.TryParse(
                GetCellValue(row.GetCell(2)),
                NumberStyles.Any,
                CultureInfo.GetCultureInfo("de-DE"),
                out var grade
            );

            var marked = GetCellValue(row.GetCell(6))
                .Trim()
                .Equals("ja", StringComparison.OrdinalIgnoreCase);

            var exam = new ExamDto(
                sheet.SheetName,
                grade,
                GetCellValue(row.GetCell(3)),
                GetCellValue(row.GetCell(4)),
                GetCellValue(row.GetCell(5)),
                marked
            );

            var key = $"{firstName}|{lastName}|{club}";

            if (lookup.TryGetValue(key, out var existing))
            {
                // ⚠️ record ist immutable → neue Instanz bauen
                var updated = existing with
                {
                    Exam = existing.Exam.Append(exam).ToList()
                };

                // ersetzen in Liste + lookup
                var index = examResults.IndexOf(existing);
                examResults[index] = updated;
                lookup[key] = updated;
            }
            else
            {
                var newEntry = new ExamResultDto(
                    firstName,
                    lastName,
                    [exam],
                    club
                );

                examResults.Add(newEntry);
                lookup[key] = newEntry;
            }
        }
    }
    
    private static List<Student> ParseStudentSheet(ISheet? sheet)
    {
        var students = new List<Student>();
        
        if (sheet is null) return students;
        
        // Start bei Zeile 1 (Zeile 0 ist Header)
        for (var rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null || IsRowEmpty(row)) continue;
            
            students.Add(new Student
            {
                FirstName = GetCellValue(row.GetCell(0)),
                LastName = GetCellValue(row.GetCell(1)),
                DateOfBirth = GetCellValueAsDate(row.GetCell(2)) ?? DateTime.Now,
                Club = GetCellValue(row.GetCell(3))
            });
        }
        
        return students;
    }
    
    private static List<Exam> ParseExamSheet(ISheet? sheet)
    {
        var exams = new List<Exam>();
        
        if (sheet is null) return exams;
        
        for (var rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null || IsRowEmpty(row)) continue;
            
            exams.Add(new Exam
            {
                Name = GetCellValue(row.GetCell(0)),
                Description = GetCellValue(row.GetCell(1))
            });
        }
        
        return exams;
    }
  
    private static string GetCellValue(ICell? cell)
    {
        if (cell is null) return string.Empty;
        
        return cell.CellType switch
        {
            CellType.String => cell.StringCellValue?.Trim() ?? string.Empty,
            CellType.Numeric => cell.NumericCellValue.ToString(new CultureInfo("de-DE")),
            CellType.Boolean => cell.BooleanCellValue.ToString(),
            CellType.Formula => cell.StringCellValue?.Trim() ?? string.Empty,
            _ => string.Empty
        };
    }
    
    private static DateTime? GetCellValueAsDate(ICell? cell)
    {
        if (cell is null) return null;
        
        try
        {
            if (cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
            {
                return cell.DateCellValue;
            }
            
            if (cell.CellType == CellType.String)
            {
                var value = cell.StringCellValue?.Trim();
                if (DateTime.TryParse(value, out var date))
                {
                    return date;
                }
            }
        }
        catch
        {
            return DateTime.Now;
        }
        
        return null;
    }
    
    private static bool IsRowEmpty(IRow row)
    {
        for (int i = row.FirstCellNum; i < row.LastCellNum; i++)
        {
            var cell = row.GetCell(i);
            if (cell != null && cell.CellType != CellType.Blank && 
                !string.IsNullOrWhiteSpace(GetCellValue(cell)))
            {
                return false;
            }
        }
        return true;
    }
}