using Domain.Entity;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Application.Services;

public class ExcelGradesExportService
{

    public static byte[] GetFile(Course course, List<ExamResult> examResults, string examiner)
    {
        // Font Suche deaktivieren für WASM
        Environment.SetEnvironmentVariable("NPOI_FONT_PATH", "");
        
        var workbook = new XSSFWorkbook();
        
        foreach (var exam in course.Exams)
        {
            var sheet = workbook.CreateSheet(exam.Name);
            SetupHeaders(sheet);
            FillSheet(sheet, examResults.Where(r => r.Exam.Id == exam.Id).ToList(), examiner);
        }
        
        using var stream = new MemoryStream();
        workbook.Write(stream, true);
        return stream.ToArray();
    }

    private static void SetupHeaders(ISheet sheet)
    {
        var headers = new[] { "Vorname", "Nachname", "Note", "Tendenz", "Prüfer", "Bemerkung", "Markiert", "Verein" };
        var workbook = sheet.Workbook;
        var headerStyle = workbook.CreateCellStyle();

        headerStyle.FillForegroundColor = IndexedColors.CornflowerBlue.Index;
        headerStyle.FillPattern = FillPattern.SolidForeground;

        headerStyle.BorderTop = BorderStyle.Thin;
        headerStyle.BorderBottom = BorderStyle.Thin;
        headerStyle.BorderLeft = BorderStyle.Thin;
        headerStyle.BorderRight = BorderStyle.Thin;
        
        headerStyle.Alignment = HorizontalAlignment.Center;
        headerStyle.VerticalAlignment = VerticalAlignment.Center;
        
        var headerRow = sheet.CreateRow(0);
        for (var i = 0; i < headers.Length; i++)
        {
            var cell = headerRow.CreateCell(i);
            cell.SetCellValue(headers[i]);
            cell.CellStyle = headerStyle;
            sheet.SetColumnWidth(i, (headers[i].Length + 2) * 256);
        }

        sheet.CreateFreezePane(0, 1);
    }

    private static void FillSheet(ISheet sheet, List<ExamResult> examResults, string examiner)
    {
        var workbook = sheet.Workbook;
        var defaultStyle = workbook.CreateCellStyle();
        defaultStyle.Alignment = HorizontalAlignment.Left;
        defaultStyle.VerticalAlignment = VerticalAlignment.Center;

        var rowIndex = 1;

        foreach (var examGroup in examResults.GroupBy(r => r.Exam.Id).OrderBy(g => g.First().Exam.Name))
        {
            foreach (var result in examGroup.OrderBy(r => r.Student.LastName).ThenBy(r => r.Student.FirstName))
            {
                var row = sheet.CreateRow(rowIndex++);
                
                var cellFirstname = row.CreateCell(0);
                cellFirstname.SetCellValue(result.Student.FirstName);
                cellFirstname.CellStyle = defaultStyle;

                var cellLastName = row.CreateCell(1);
                cellLastName.SetCellValue(result.Student.LastName);
                cellLastName.CellStyle = defaultStyle;
                
                var cellGrade = row.CreateCell(2);
                if (result.Grade > 0.0f)
                {
                    cellGrade.SetCellValue(result.Grade);
                    cellGrade.CellStyle = defaultStyle;
                    cellGrade.SetCellType(CellType.Numeric);
                }

                var cellTendency= row.CreateCell(3);
                cellTendency.SetCellValue(result.Tendency);
                cellTendency.CellStyle = defaultStyle;
                
                var cellExaminer = row.CreateCell(4);
                cellExaminer.SetCellValue(examiner);
                cellExaminer.CellStyle = defaultStyle;
                
                var cellComment = row.CreateCell(5);
                cellComment.SetCellValue(result.Comment ?? string.Empty);
                cellComment.CellStyle = defaultStyle;
                
                var cellIsMarked = row.CreateCell(6);
                cellIsMarked.SetCellValue(result.IsMarked ? "ja" : " ");
                cellIsMarked.CellStyle = defaultStyle;
                
                var cellClub = row.CreateCell(7);
                cellClub.SetCellValue(result.Student.Club);
                cellClub.CellStyle = defaultStyle;
            }
        }
    }
}