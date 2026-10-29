using System.Globalization;

namespace ScheduleParser;

internal static class Program
{
    private static int Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        ShowWelcomeTemplate();
        return StartProgram(args);
    }

    private static int StartProgram(string[] args)
    {
        try
        {
            var options = ReadOptions(args);
            if (!Directory.Exists(options.InputDirectory))
            {
                Console.Error.WriteLine($"Папка не найдена: {options.InputDirectory}");
                return 1;
            }

            if (string.IsNullOrWhiteSpace(options.Teacher))
            {
                Console.Error.WriteLine("Не указана фамилия или ФИО преподавателя.");
                return 1;
            }

            var scan = new ScheduleScanner(new ExcelWorkbookReader(), new ScheduleExtractor())
                .Scan(options.InputDirectory, options.Teacher, options.EffectiveKeyword, Console.WriteLine);

            foreach (var warning in scan.Warnings)
                Console.Error.WriteLine($"Предупреждение: {warning}");

            Directory.CreateDirectory(options.OutputDirectory);
            var outputPath = Path.Combine(options.OutputDirectory,
                $"output{DateTime.Now.ToString("dd_MM_yy_HH_mm_ss", CultureInfo.InvariantCulture)}.xlsx");
            new ExcelReportWriter().Write(outputPath, scan.Entries);
            if (scan.Entries.Count == 0)
                Console.Error.WriteLine("ВНИМАНИЕ: занятий не найдено. Проверьте ФИО, тип файлов и предупреждения выше.");

            Console.WriteLine($"Найдено файлов: {scan.MatchedFiles}; обработано: {scan.ProcessedFiles}; занятий: {scan.Entries.Count}.");
            Console.WriteLine($"Результаты сохранены: {outputPath}");
            return scan.Warnings.Count > 0 ? 2 : 0;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
        {
            Console.Error.WriteLine($"Ошибка: {ex.Message}");
            return 1;
        }
    }

    static void ShowWelcomeTemplate()
    {
        Console.Title = "Парсер расписания - Alexey Dmitriev";

        Console.ForegroundColor = ConsoleColor.White;
        Console.WriteLine(new string('=', 60));
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("            ПАРСЕР РАСПИСАНИЯ");
        Console.ForegroundColor = ConsoleColor.White;
        Console.WriteLine(new string('=', 60));
        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine("  Версия 2.0 | Считывание информации из excel");
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("  Разработчик: Alexey Dmitriev");
        Console.WriteLine("  GitHub: https://github.com/cptfrosty");
        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine(new string('-', 60));
        Console.ResetColor();
        Console.WriteLine();
    }

    private static ParserOptions ReadOptions(string[] args)
    {
        if (args.Length == 4)
            return new ParserOptions(args[0], args[1], args[2], args[3]);

        if (args.Length != 0)
            throw new ArgumentException("Укажите 4 аргумента: <папка> <преподаватель> <ключевое слово> <папка результата>.");

        Console.Write("(1/4) Путь к папке с расписанием: ");
        var input = Console.ReadLine()?.Trim().Trim('"') ?? "";
        Console.Write("(2/4) Фамилия или ФИО преподавателя: ");
        var teacher = Console.ReadLine()?.Trim() ?? "";
        Console.Write("(3/4) Ключевое слово файла [лист замен]: ");
        var keyword = Console.ReadLine()?.Trim() ?? "";
        Console.Write("(4/4) Папка для результата: ");
        var output = Console.ReadLine()?.Trim().Trim('"') ?? "";
        return new ParserOptions(input, teacher, keyword, output);
    }
}

internal sealed record ParserOptions(string InputDirectory, string Teacher, string FileKeyword, string OutputDirectory)
{
    public string EffectiveKeyword => string.IsNullOrWhiteSpace(FileKeyword) ? "лист замен" : FileKeyword;
}
