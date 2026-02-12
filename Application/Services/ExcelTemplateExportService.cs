using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Application.Services;

public class ExcelTemplateExportService
{
    public byte[] GenerateTemplate()
    {
        // Font Suche deaktivieren für WASM
        Environment.SetEnvironmentVariable("NPOI_FONT_PATH", "");
        
        var workbook = new XSSFWorkbook();
        
        // Teilnehmer Worksheet
        var students = workbook.CreateSheet("Teilnehmer");
        SetupStudentSheet(students);
        
        // Prüfungen Worksheet
        var exams = workbook.CreateSheet("Prüfungen"); SetupExamSheet(exams);
        
        // Workbook als byte[] zurückgeben
        using var stream = new MemoryStream();
        workbook.Write(stream, true);
        return stream.ToArray();
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