namespace ScheduleParser;

public sealed record ScanResult(
    IReadOnlyList<ScheduleEntry> Entries,
    IReadOnlyList<string> Warnings,
    int MatchedFiles,
    int ProcessedFiles);

public sealed class ScheduleScanner(IWorkbookReader reader, ScheduleExtractor extractor)
{
    public ScanResult Scan(string directory, string teacher, string keyword = "лист замен", Action<string>? progress = null)
    {
        if (!Directory.Exists(directory))
            throw new DirectoryNotFoundException($"Папка не найдена: {directory}");

        if (string.IsNullOrWhiteSpace(teacher))
            throw new ArgumentException("Не указана фамилия или ФИО преподавателя.", nameof(teacher));

        keyword = string.IsNullOrWhiteSpace(keyword) ? "лист замен" : keyword;
        var normalizedKeyword = NormalizeKeyword(keyword);
        var files = Directory.EnumerateFiles(directory, "*", SearchOption.AllDirectories)
            .Where(path => Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase)
                || Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
            .Where(path => NormalizeKeyword(Path.GetFileNameWithoutExtension(path)).Contains(normalizedKeyword, StringComparison.Ordinal))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var entries = new List<ScheduleEntry>();
        var warnings = new List<string>();
        var processed = 0;

        for (var index = 0; index < files.Length; index++)
        {
            var path = files[index];
            progress?.Invoke($"[{index + 1}/{files.Length}] Файл: {path}");
            try
            {
                var sheets = reader.Read(path);
                var anyUsableSheet = false;
                var foundInFile = 0;
                foreach (var sheet in sheets)
                {
                    if (!extractor.TryExtract(sheet, path, teacher, out var found, out var warning))
                    {
                        if (!string.IsNullOrEmpty(warning))
                            warnings.Add(warning);
                        continue;
                    }

                    anyUsableSheet = true;
                    entries.AddRange(found);
                    foundInFile += found.Count;
                    progress?.Invoke($"    Лист: {sheet.Name}; найдено занятий: {found.Count}");
                    foreach (var lesson in found)
                        progress?.Invoke($"      {lesson.Date:dd.MM.yyyy} | {lesson.Group} | пара {lesson.LessonNumber} | {lesson.Subject} | {lesson.Teacher} | {lesson.Classroom}");
                    if (!string.IsNullOrEmpty(warning))
                        warnings.Add(warning);
                }

                if (anyUsableSheet)
                {
                    processed++;
                }
                else if (sheets.Count == 0)
                {
                    warnings.Add($"{path}: файл не содержит листов.");
                }
                else
                {
                    warnings.Add($"{path}: не удалось распознать таблицу.");
                }

                progress?.Invoke(
                    $"    Итог файла: {foundInFile} занят(ий)" +
                    (anyUsableSheet ? "" : "; не удалось распознать таблицу")
                );
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException
                or NotSupportedException or System.IO.InvalidDataException)
            {
                warnings.Add($"{path}: ошибка чтения ({ex.Message}).");
                progress?.Invoke($"    ОШИБКА ЧТЕНИЯ: {ex.Message}");
            }
        }

        // Сохранять каждый урок один раз, если файлы за один и тот же день
        // публикуются как в исходном виде, так и в виде копии «(1)».
        // Группа и подгруппа являются частью ключа.
        var unique = entries.DistinctBy(item => new
        {
            item.Date,
            Pair = item.LessonNumber.ToUpperInvariant(),
            Group = item.Group.ToUpperInvariant(),
            Subgroup = item.Subgroup.ToUpperInvariant(),
            Subject = item.Subject.ToUpperInvariant(),
            Teacher = item.Teacher.ToUpperInvariant(),
            Classroom = item.Classroom.ToUpperInvariant(),
            Territory = item.Territory.ToUpperInvariant()
        });

        var sorted = unique.OrderBy(x => x.Date)
            .ThenBy(x => ParseLessonOrder(x.LessonNumber))
            .ThenBy(x => x.Group, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(x => x.Subgroup, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();

        return new ScanResult(sorted, warnings, files.Length, processed);
    }

    private static int ParseLessonOrder(string lesson)
    {
        var number = lesson.Split('-', '–', '/', ',')[0].Trim();
        return int.TryParse(number, out var parsed) ? parsed : int.MaxValue;
    }

    private static string NormalizeKeyword(string keyword) =>
        ScheduleExtractor.Normalize(keyword).Replace('ё', 'е').ToLowerInvariant();
}
