using System.Globalization;
using System.Text.RegularExpressions;

namespace ScheduleParser;

public sealed class ScheduleExtractor
{
    private static readonly Regex DatePattern = new(@"(?<!\d)(\d{1,2}\.\d{1,2}\.\d{2,4})(?!\d)", RegexOptions.Compiled);
    private static readonly Regex GroupPattern = new(@"^\d{1,2}\s*[-–]\s*\d{2}\b", RegexOptions.Compiled);
    private static readonly Regex LessonPattern = new(@"^\d{1,2}(?:\s*[-–/,]\s*\d{1,2})?$", RegexOptions.Compiled);
    private static readonly string[] DateFormats = ["d.M.yyyy", "dd.MM.yyyy", "d.M.yy", "dd.MM.yy"];

    public bool TryExtract(ScheduleSheet sheet, string sourceFile, string teacherQuery,
        out IReadOnlyList<ScheduleEntry> entries, out string warning)
    {
        entries = [];
        warning = string.Empty;

        if (!TryFindHeader(sheet, out var header))
        {
            warning = $"{sourceFile}: лист «{sheet.Name}» не содержит заголовков таблицы замен.";
            return false;
        }

        if (!TryResolveDate(sheet, sourceFile, out var date, out var dateWarning))
        {
            warning = $"{sourceFile}: не удалось определить дату из таблицы и имени файла.";
            return false;
        }

        warning = dateWarning;
        var result = new List<ScheduleEntry>();
        var group = string.Empty;

        for (var row = header.RowIndex + 1; row < sheet.Rows.Count; row++)
        {
            var pair = Normalize(sheet.Cell(row, header.Pair));
            if (GroupPattern.IsMatch(pair))
            {
                group = pair;
                continue;
            }

            if (!LessonPattern.IsMatch(pair))
                continue;

            var teacher = Normalize(sheet.Cell(row, header.Teacher));
            if (!TeacherMatches(teacher, teacherQuery))
                continue;

            result.Add(new ScheduleEntry(
                date,
                pair,
                group,
                Normalize(sheet.Cell(row, header.Subgroup)),
                Normalize(sheet.Cell(row, header.Subject)),
                teacher,
                Normalize(sheet.Cell(row, header.Classroom)),
                Normalize(sheet.Cell(row, header.Territory)),
                Path.GetFileName(sourceFile),
                sheet.Name,
                row + 1));
        }

        entries = result;
        return true;
    }

    public static bool TeacherMatches(string value, string query)
    {
        // The first token (surname) must match completely; subsequent tokens
        // may be prefixes (e.g. "Дмитриев А" for "Дмитриев Алексей Андреевич").
        // Initials with dots and names separated by '/' are supported.
        static string[] Tokens(string text) => Regex.Matches(
                Normalize(text).Replace('ё', 'е').Replace('Ё', 'Е'), @"[\p{L}]+")
            .Select(match => match.Value)
            .ToArray();

        var requested = Tokens(query);
        if (requested.Length == 0)
            return false;

        foreach (var candidate in value.Split('/', ';', '\n'))
        {
            var actual = Tokens(candidate);
            if (actual.Length < requested.Length ||
                !actual[0].Equals(requested[0], StringComparison.OrdinalIgnoreCase))
                continue;

            if (requested.Skip(1).Select((part, i) =>
                    actual[i + 1].StartsWith(part, StringComparison.OrdinalIgnoreCase)).All(match => match))
                return true;
        }

        return false;
    }

    public static string Normalize(string? value) => Regex.Replace(value ?? "", @"\s+", " ").Trim();

    private static bool TryFindHeader(ScheduleSheet sheet, out Header header)
    {
        for (var r = 0; r < Math.Min(50, sheet.Rows.Count); r++)
        {
            var values = sheet.Rows[r].Select(x => Normalize(x).ToLowerInvariant()).ToArray();
            var pair = FindColumn(values, "номер пары", "№ пары");
            var subject = FindColumn(values, "дисциплин");
            var teacher = FindColumn(values, "преподавател");
            if (pair < 0 || subject < 0 || teacher < 0)
                continue;

            header = new Header(r, pair, subject, teacher,
                FindColumn(values, "аудитор"), FindColumn(values, "территор"), FindColumn(values, "подгруп"));
            return true;
        }

        header = default;
        return false;
    }

    private static int FindColumn(string[] values, params string[] labels) =>
        Array.FindIndex(values, value => labels.Any(label => value.Contains(label, StringComparison.Ordinal)));

    public static bool TryResolveDate(ScheduleSheet sheet, string fileName, out DateOnly date, out string warning)
    {
        warning = string.Empty;
        DateOnly? titleDate = null;
        foreach (var row in sheet.Rows.Take(15))
        {
            foreach (var cell in row)
            {
                if (TryFindDate(cell, out var parsed))
                {
                    titleDate = parsed;
                    break;
                }
            }

            if (titleDate.HasValue)
                break;
        }

        var hasFileDate = TryFindDate(Path.GetFileNameWithoutExtension(fileName), out var fileDate);
        if (titleDate.HasValue)
        {
            date = titleDate.Value;
            if (hasFileDate && fileDate != date)
                warning = $"{Path.GetFileName(fileName)}: дата в заголовке ({date:dd.MM.yyyy}) не совпадает с именем файла ({fileDate:dd.MM.yyyy}); выбрана дата из Excel.";
            return true;
        }

        date = fileDate;
        return hasFileDate;
    }

    private static bool TryFindDate(string? text, out DateOnly date)
    {
        foreach (Match match in DatePattern.Matches(text ?? ""))
        {
            if (DateOnly.TryParseExact(match.Value, DateFormats, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out date))
                return true;
        }

        date = default;
        return false;
    }

    private readonly record struct Header(int RowIndex, int Pair, int Subject, int Teacher,
        int Classroom, int Territory, int Subgroup);
}
