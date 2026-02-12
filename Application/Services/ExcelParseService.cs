using System.Globalization;
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