using ClosedXML.Excel;

namespace ScheduleParser;

public sealed class ExcelReportWriter
{
    public void Write(string outputPath, IReadOnlyList<ScheduleEntry> entries)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ArgumentNullException.ThrowIfNull(entries);

        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Результаты");
        string[] headers = ["Дата", "Пара", "Учебная группа", "Подгруппа", "Дисциплина",
            "Преподаватель", "Аудитория", "Территория", "Источник", "Лист", "Строка Excel"];

        for (var c = 0; c < headers.Length; c++)
            worksheet.Cell(1, c + 1).Value = headers[c];

        for (var i = 0; i < entries.Count; i++)
        {
            var entry = entries[i];
            var row = i + 2;
            worksheet.Cell(row, 1).Value = entry.Date.ToDateTime(TimeOnly.MinValue);
            worksheet.Cell(row, 2).Value = entry.LessonNumber;
            worksheet.Cell(row, 3).Value = entry.Group;
            worksheet.Cell(row, 4).Value = entry.Subgroup;
            worksheet.Cell(row, 5).Value = entry.Subject;
            worksheet.Cell(row, 6).Value = entry.Teacher;
            worksheet.Cell(row, 7).Value = entry.Classroom;
            worksheet.Cell(row, 8).Value = entry.Territory;
            worksheet.Cell(row, 9).Value = entry.SourceFile;
            worksheet.Cell(row, 10).Value = entry.SheetName;
            worksheet.Cell(row, 11).Value = entry.ExcelRow;
        }

        var header = worksheet.Range(1, 1, 1, headers.Length);
        header.Style.Font.Bold = true;
        header.Style.Font.FontColor = XLColor.White;
        header.Style.Fill.BackgroundColor = XLColor.DarkBlue;
        worksheet.Column(1).Style.DateFormat.Format = "dd.MM.yyyy";
        worksheet.SheetView.FreezeRows(1);
        worksheet.Columns(1, headers.Length).AdjustToContents();
        worksheet.Column(5).Width = Math.Min(worksheet.Column(5).Width, 60);
        worksheet.Column(9).Width = Math.Min(worksheet.Column(9).Width, 45);
        worksheet.Range(1, 1, Math.Max(2, entries.Count + 1), headers.Length).SetAutoFilter();
        workbook.SaveAs(outputPath);
    }
}
