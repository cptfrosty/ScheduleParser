using ScheduleParser;
using Xunit;

namespace ScheduleParser.Tests;

public sealed class ScheduleExtractorTests
{
    private readonly ScheduleExtractor _parser = new();

    [Fact]
    public void StandardTemplate_FindsLessonAndItsGroup()
    {
        var sheet = Sheet(
            ["Лист замен на 02.09.2026 СРЕДА"],
            ["Учебная группа"],
            ["Номер пары", "Дисциплина", "", "", "Преподаватель", "", "Аудитория", "Территория"],
            ["1-26 ИСП-1"],
            ["1", "Математика", "", "", "Сидорова Елена Петровна", "", "203", "1 корпус"],
            ["2", "Программирование", "", "", "Иванов Иван Иванович", "", "215", "1 корпус"]);

        var ok = _parser.TryExtract(sheet, "02.09.2026 (лист замен).xlsx", "Иванов", out var lessons, out var warning);

        Assert.True(ok);
        Assert.Empty(warning);
        var lesson = Assert.Single(lessons);
        Assert.Equal(new DateOnly(2026, 9, 2), lesson.Date);
        Assert.Equal("1-26 ИСП-1", lesson.Group);
        Assert.Equal("2", lesson.LessonNumber);
        Assert.Equal("Программирование", lesson.Subject);
        Assert.Equal("215", lesson.Classroom);
    }

    [Fact]
    public void FirstSeptemberTemplate_ResolvesShiftedColumnsAndDateInMetadata()
    {
        var sheet = Sheet(
            ["Расписание"],
            ["Параметры:", "", "", "Дата расписания: 01.09.2026 0:00:00"],
            ["Учебная группа"],
            ["Номер пары", "Подгруппа", "", "", "Дисциплина", "", "", "Преподаватель", "Аудитория", "Территория"],
            ["1-26 ИСП-2"],
            ["3", "2", "", "", "Базы данных", "", "", "Иванов И.И.", "308", "1 корпус"]);

        var ok = _parser.TryExtract(sheet, "01.09.2026 Лист замен.xlsx", "Иванов", out var lessons, out _);

        Assert.True(ok);
        var lesson = Assert.Single(lessons);
        Assert.Equal(new DateOnly(2026, 9, 1), lesson.Date);
        Assert.Equal("Базы данных", lesson.Subject);
        Assert.Equal("2", lesson.Subgroup);
        Assert.Equal("308", lesson.Classroom);
    }

    [Fact]
    public void OctoberTemplate_FindsClassroomInColumnF()
    {
        var sheet = Sheet(
            ["Лист замен на 15.10.2026 ЧЕТВЕРГ"],
            ["Номер пары", "Дисциплина", "", "", "Преподаватель", "Аудитория", "", "Территория"],
            ["1-26 ИСП-3"],
            ["4", "Информатика", "", "", "Иванов Иван", "2.5ВЦ", "", "ВЦ"]);

        Assert.True(_parser.TryExtract(sheet, "15.10.2026 лист замен.xlsx", "Иванов", out var lessons, out _));
        Assert.Equal("2.5ВЦ", Assert.Single(lessons).Classroom);
    }

    [Fact]
    public void DateInTitleWinsOverMismatchedFilename()
    {
        var sheet = Sheet(["Лист замен на 04.09.2026"],
            ["Номер пары", "Дисциплина", "Преподаватель"],
            ["1-26 ИСП-1"],
            ["1", "Тестирование", "Иванов"]);

        Assert.True(ScheduleExtractor.TryResolveDate(sheet, "04.09.2029 лист замен.xlsx", out var date, out var warning));
        Assert.Equal(new DateOnly(2026, 9, 4), date);
        Assert.Contains("не совпадает", warning);
    }

    [Theory]
    [InlineData("Дмитриев Алексей Андреевич", "Дмитриев А", true)]
    [InlineData("Дмитриев Алексей Андреевич", "Дмитриев А.А.", true)]
    [InlineData("Дмитриев Алексей Андреевич", "Дмитриев Алексей", true)]
    [InlineData("Орин Павел Дмитриевич", "Дмитриев А", false)]
    [InlineData("Дмитриева Анна", "Дмитриев А", false)]
    [InlineData("Иванов Иван", "Иванов", true)]
    [InlineData("Иванова Ирина", "Иванов", false)]
    [InlineData("Петров / Иванов", "Иванов", true)]
    [InlineData("Семёнов Семён", "Семенов", true)]
    [InlineData("Иванович Дмитрий", "Иванов", false)]
    public void MatchesWholeTeacherNameTokens(string teacher, string query, bool expected)
    {
        Assert.Equal(expected, ScheduleExtractor.TeacherMatches(teacher, query));
    }

    [Fact]
    public void PartialTeacherName_ExtractsActualLesson()
    {
        var sheet = Sheet(
            ["Лист замен на 15.10.2026 ЧЕТВЕРГ"],
            ["Учебная группа"],
            ["Номер пары", "Дисциплина", "", "", "Преподаватель", "Аудитория", "", "Территория"],
            ["3-24 ИСП-1"],
            ["2", "Технология разработки ПО", "", "", "Дмитриев Алексей Андреевич", "216", "", "1 корпус"]);
        Assert.True(_parser.TryExtract(sheet, "15.10.2026 Лист замен.xlsx", "Дмитриев А", out var found, out _));
        Assert.Single(found);
    }

    [Fact]
    public void WithoutTeacherMatches_ReturnsEmptyCollection()
    {
        var sheet = Sheet(["Лист замен на 05.09.2026"],
            ["Номер пары", "Дисциплина", "Преподаватель"],
            ["1-26 ИСП-1"], ["1", "История", "Другой Преподаватель"]);

        Assert.True(_parser.TryExtract(sheet, "05.09.2026 лист замен.xlsx", "Иванов", out var lessons, out _));
        Assert.Empty(lessons);
    }

    private static ScheduleSheet Sheet(params string[][] rows) => new("Лист_1", rows);
}
