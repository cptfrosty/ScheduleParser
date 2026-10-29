using ScheduleParser;
using Xunit;

namespace ScheduleParser.Tests;

public sealed class ScheduleScannerTests
{
    [Fact]
    public void SearchesBothExcelExtensionsAndDeduplicatesCopies()
    {
        var folder = Path.Combine(Path.GetTempPath(), "ScheduleParserTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            File.WriteAllText(Path.Combine(folder, "14.09.2026 (лист замен).xlsx"), "fake");
            File.WriteAllText(Path.Combine(folder, "14.09.2026 (лист замен)(1).xlsx"), "fake");
            File.WriteAllText(Path.Combine(folder, "14.09.2026 (лист замен).xls"), "fake");
            File.WriteAllText(Path.Combine(folder, "14.09.2026 (занятость).xlsx"), "fake");

            var scanner = new ScheduleScanner(new StubWorkbookReader(), new ScheduleExtractor());
            var result = scanner.Scan(folder, "Иванов", "лист замен");

            Assert.Equal(3, result.MatchedFiles);
            Assert.Equal(3, result.ProcessedFiles);
            Assert.Single(result.Entries);
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void ExcelReportPreservesFirstRowAndSemicolonInSubject()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ScheduleParser_{Guid.NewGuid():N}.xlsx");
        try
        {
            var lesson = new ScheduleEntry(new DateOnly(2026, 10, 15), "1", "1-26 ИСП-1", "",
                "Анализ; проектирование", "Иванов Иван Иванович", "216", "ВЦ", "test.xls", "Лист_1", 6);
            new ExcelReportWriter().Write(path, [lesson]);
            var sheet = Assert.Single(new ExcelWorkbookReader().Read(path));
            Assert.Equal("Дата", sheet.Cell(0, 0));
            Assert.Equal("Анализ; проектирование", sheet.Cell(1, 4));
            Assert.Equal("Иванов Иван Иванович", sheet.Cell(1, 5));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ExcelReportCreatesHeadersWithNoMatches()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ScheduleParser_{Guid.NewGuid():N}.xlsx");
        try
        {
            new ExcelReportWriter().Write(path, []);
            var sheet = Assert.Single(new ExcelWorkbookReader().Read(path));
            Assert.Equal("Дата", sheet.Cell(0, 0));
            Assert.Equal("Строка Excel", sheet.Cell(0, 10));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private sealed class StubWorkbookReader : IWorkbookReader
    {
        public IReadOnlyList<ScheduleSheet> Read(string path) =>
        [
            new ScheduleSheet("Лист_1", new[]
            {
                new[] { "Лист замен на 14.09.2026" },
                new[] { "Номер пары", "Дисциплина", "Преподаватель" },
                new[] { "1-26 ИСП-1" },
                new[] { "2", "Базы данных", "Иванов Иван Иванович" }
            })
        ];
    }
}
