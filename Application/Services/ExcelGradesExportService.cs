using Domain.Entity;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Application.Services;

public class ExcelGradesExportService
{
    public byte[] GetExcelFile(Course course, List<ExamResult> examResults, string examiner)
    {
        // Font Suche deaktivieren für WASM
        Environment.SetEnvironmentVariable("NPOI_FONT_PATH", "");
        
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Noten");

        SetupHeaders(sheet);
        FillSheet(sheet, course, examResults, examiner);

        using var stream = new MemoryStream();
        workbook.Write(stream, true);
        return stream.ToArray();
    }

    private void SetupHeaders(ISheet sheet)
    {
        var headers = new[] { "Prüfungsname", "Beschreibung", "Teilnehmer", "Verein", "Note", "Prüfer", "Bemerkung" };
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
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = headerRow.CreateCell(i);
            cell.SetCellValue(headers[i]);
            cell.CellStyle = headerStyle;
            sheet.SetColumnWidth(i, (headers[i].Length + 2) * 256);
        }

        sheet.CreateFreezePane(0, 1);
    }

    private void FillSheet(ISheet sheet, Course course, List<ExamResult> examResults, string examiner)
    {
        var workbook = sheet.Workbook;
        var defaultStyle = workbook.CreateCellStyle();
        defaultStyle.Alignment = HorizontalAlignment.Left;
        defaultStyle.VerticalAlignment = VerticalAlignment.Center;

        int rowIndex = 1;

        foreach (var examGroup in examResults.GroupBy(r => r.Exam.Id).OrderBy(g => g.First().Exam.Name))
        {
            var exam = examGroup.First().Exam;
            foreach (var result in examGroup.OrderBy(r => r.Student.LastName).ThenBy(r => r.Student.FirstName))
            {
                var row = sheet.CreateRow(rowIndex++);

                var cellExamName = row.CreateCell(0);
                cellExamName.SetCellValue(exam.Name);
                cellExamName.CellStyle = defaultStyle;

                var cellDescription = row.CreateCell(1);
                cellDescription.SetCellValue(exam.Description);
                cellDescription.CellStyle = defaultStyle;

                var cellStudent = row.CreateCell(2);
                cellStudent.SetCellValue($"{result.Student.FirstName} {result.Student.LastName}");
                cellStudent.CellStyle = defaultStyle;

                var cellClub = row.CreateCell(3);
                cellClub.SetCellValue(result.Student.Club);
                cellClub.CellStyle = defaultStyle;

                var cellGrade = row.CreateCell(4);
                cellGrade.SetCellValue($"{result.Grade:0.0}{result.Tendency}");
                cellGrade.CellStyle = defaultStyle;

                var cellExaminer = row.CreateCell(5);
                cellExaminer.SetCellValue(examiner);
                cellExaminer.CellStyle = defaultStyle;

                var cellComment = row.CreateCell(6);
                cellComment.SetCellValue(result.Comment ?? string.Empty);
                cellComment.CellStyle = defaultStyle;
            }
        }
    }
}