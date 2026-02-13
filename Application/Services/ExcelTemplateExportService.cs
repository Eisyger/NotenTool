using Domain.Entity;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Application.Services;

public class ExcelTemplateExportService
{
    public byte[] GenerateEmptyTemplate()
    {
        // Font Suche deaktivieren für WASM
        Environment.SetEnvironmentVariable("NPOI_FONT_PATH", "");
        
        var workbook = new XSSFWorkbook();
        
        // Teilnehmer Worksheet
        var students = workbook.CreateSheet("Teilnehmer");
        SetupStudentSheet(students);
        
        // Prüfungen Worksheet
        var exams = workbook.CreateSheet("Prüfungen"); 
        SetupExamSheet(exams);
        
        // Workbook als byte[] zurückgeben
        using var stream = new MemoryStream();
        workbook.Write(stream, true);
        return stream.ToArray();
    }

    public byte[] GenerateExamOrderTemplate(Dictionary<string, List<Student>> students)
    {
        // Font Suche deaktivieren für WASM
        Environment.SetEnvironmentVariable("NPOI_FONT_PATH", "");
        
        var workbook = new XSSFWorkbook();

        foreach (var key in students.Keys)
        {
            var sheet = workbook.CreateSheet(key); 
            SetupExamOrderSheet(sheet);
            SetExamOrderData(sheet, students[key]);
        }
        using var stream = new MemoryStream();
        workbook.Write(stream, true);
        return stream.ToArray();
    }

    private void SetExamOrderData(ISheet sheet, List<Student> students)
    {  
        var workbook = sheet.Workbook;
        var defaultStyle = workbook.CreateCellStyle();
        defaultStyle.Alignment = HorizontalAlignment.Left;
        defaultStyle.VerticalAlignment = VerticalAlignment.Center;
        
        int rowIndex = 1;

        foreach (var student in students)
        {
            var row = sheet.CreateRow(rowIndex++);

            var cellExamName = row.CreateCell(0);
            cellExamName.SetCellValue(rowIndex-1);
            cellExamName.CellStyle = defaultStyle;

            var cellDescription = row.CreateCell(1);
            cellDescription.SetCellValue(student.FirstName);
            cellDescription.CellStyle = defaultStyle;

            var cellStudent = row.CreateCell(2);
            cellStudent.SetCellValue(student.LastName);
            cellStudent.CellStyle = defaultStyle;

            var cellClub = row.CreateCell(3);
            cellClub.SetCellValue(student.DateOfBirth.ToString("dd.MM.yyyy"));
            cellClub.CellStyle = defaultStyle;

            var cellGrade = row.CreateCell(4);
            cellGrade.SetCellValue(student.Club);
            cellGrade.CellStyle = defaultStyle;
        }
    }
    

    private void SetupExamOrderSheet(ISheet sheet)
    {
        var headers = new[] {"Nr.", "Vorname", "Nachname", "Geburtsdatum", "Verein" };
        SetHeaders(sheet, headers);
    }
    
    private void SetupStudentSheet(ISheet sheet)
    {
        var headers = new[] { "Vorname", "Nachname", "Geburtsdatum", "Verein" };
        SetHeaders(sheet, headers);
    }
    
    private void SetupExamSheet(ISheet sheet)
    {
        var headers = new[] { "Prüfungsname", "Beschreibung" };
        SetHeaders(sheet, headers);
    }
    
    private void SetHeaders(ISheet sheet, string[] headers)
    {
        var headerRow = sheet.CreateRow(0);
        var headerStyle = CreateHeaderStyle(sheet.Workbook);
        
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = headerRow.CreateCell(i);
            cell.SetCellValue(headers[i]);
            cell.CellStyle = headerStyle;
            
            int width = (headers[i].Length + 2) * 256; // +2 für Padding
            sheet.SetColumnWidth(i, width);
        }
        
        // Freeze header row
        sheet.CreateFreezePane(0, 1);
    }
    
    private ICellStyle CreateHeaderStyle(IWorkbook workbook)
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