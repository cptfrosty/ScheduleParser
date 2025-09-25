using System;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using ScheduleParser;

public struct InputDataUser 
{
    public string FolderPath;
    public string SearchSecondName;
    public string FileKeyword;
    public string OutputPath;
}

class Program
{
    private static List<string> Result = new List<string>();
    private static List<Row> Rows = new List<Row>();
    private static InputDataUser _inputDataUser = new InputDataUser();
    
    static void Main(string[] args)
    {
        ShowWelcomeTemplate();
        InitializationExcelPackage();
        InputData();
        Run();
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
        Console.WriteLine("  Версия 1.0 | Считывание информации из excel");
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("  Разработчик: Alexey Dmitriev");
        Console.WriteLine("  GitHub: https://github.com/cptfrosty");
        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine(new string('-', 60));
        Console.ResetColor();
        Console.WriteLine();
    }

    static void InitializationExcelPackage()
    {
        ExcelPackage.License.SetNonCommercialOrganization("My Noncommercial organization");
    }

    static void InputData()
    {
        Console.WriteLine("(1/4 шагов) Введите путь к папке с excel расписанием:");
        _inputDataUser.FolderPath = Console.ReadLine();

        Console.WriteLine("(2/4 шагов) Введите искомую фамилию:");
        _inputDataUser.SearchSecondName = Console.ReadLine();

        Console.WriteLine("(3/4 шагов) Введите ключевое слово в названии файла:");
        _inputDataUser.FileKeyword = Console.ReadLine();

        Console.WriteLine("(4/4 шагов) Введите путь для сохранения результатов:");
        _inputDataUser.OutputPath = Console.ReadLine() + $"\\output{DateTime.Now.ToString("dd_MM_yy_HH_mm_ss")}.xlsx";
    }

    static void Run()
    {
        try
        {
            // Получаем все xlsx файлы в папке и подпапках
            var files = Directory.GetFiles(_inputDataUser.FolderPath, "*.xlsx", SearchOption.AllDirectories)
                                .Where(f => Path.GetFileName(f).Contains(_inputDataUser.FileKeyword, StringComparison.OrdinalIgnoreCase))
                                .OrderBy(f => f);

            if (!files.Any())
            {
                Console.WriteLine("Файлы не найдены.");
                return;
            }

            foreach (var filePath in files)
            {
                ProcessExcelFile(filePath, _inputDataUser.SearchSecondName);
            }

            //Сортировка
            Rows = SortRowsToDateAndLessonNumber(Rows);

            //Конвертируем тип данных Row в стринг для экспорта в Excel
            Result = ConvertRowsToResultRows(Rows);

            // Сохраняем результаты
            SaveResultsToExcel(_inputDataUser.OutputPath);

            Console.WriteLine($"Результаты сохранены в: {_inputDataUser.OutputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка: {ex.Message}");
        }
    }



    static void ProcessExcelFile(string filePath, string searchWord)
    {
        try
        {
            var fileInfo = new FileInfo(filePath);
            using (var package = new ExcelPackage(fileInfo))
            {
                var workbook = package.Workbook;
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine($"Файл {filePath} не содержит листов.");
                    return;
                }

                var worksheet = workbook.Worksheets[0]; // Первый лист

                // Извлекаем дату из имени файла
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string dateFromFileName = ExtractDateFromFileName(fileName);

                bool foundAny = false;
                int matchCount = 0;

                // Перебираем строки (начиная со второй, предполагая что первая - заголовок)
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    // Проверяем 3-ю колонку (столбец C)
                    var cellValue = worksheet.Cells[row, 5]?.Text;
                    if (cellValue != null && cellValue.Contains(searchWord, StringComparison.OrdinalIgnoreCase))
                    {
                        if (!foundAny)
                        {
                            Console.WriteLine($"\nФайл: {fileName}");
                            Console.WriteLine($"Дата из файла: {dateFromFileName}");
                            foundAny = true;
                        }

                        matchCount++;

                        // Собираем данные из 5 столбцов
                        var rowData = new List<string>();
                        for (int col = 1; col <= 7; col++)
                        {
                            string header = worksheet.Cells[1, col]?.Text ?? $"Столбец {col}";
                            string value = worksheet.Cells[row, col]?.Text ?? "";
                            Console.Write($"{value} ");
                            rowData.Add(worksheet.Cells[row, col]?.Text ?? "");
                        }

                        // Формируем строку результата
                        string resultLine = $"\"{Path.GetFileName(filePath)}\";\"{dateFromFileName}\";{row};" +
                                          $"\"{rowData[0]}\";\"{rowData[1]}\";\"{rowData[2]}\";\"{rowData[3]}\";\"{rowData[4]}\"";

                        Result.Add(resultLine);

                        //Добавление данных
                        Row outputRow = new Row();
                        outputRow.Date = dateFromFileName;
                        outputRow.LessonNumber = rowData[0];
                        outputRow.NameLesson = rowData[1];
                        outputRow.Teacher = rowData[4];
                        outputRow.ClassroomNumber = rowData[6];

                        Rows.Add(outputRow);

                        Console.WriteLine();
                    }
                }

                if (!foundAny)
                {
                    Console.WriteLine($"В файле {filePath} совпадений не найдено.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при обработке файла {filePath}: {ex.Message}");
        }
    }

    static List<string> ConvertRowsToResultRows(List<Row> rows)
    {
        List<string> result = new List<string>();

        foreach (var row in rows)
        {
            string resultLine =
                $"\"{row.Date}\";" +
                $"\"{row.LessonNumber}\";" +
                $"\"{row.NameLesson}\";" +
                $"\"{row.Teacher}\";" +
                $"\"{row.ClassroomNumber}\"";
            result.Add(resultLine);
        }

        return result;
    }

    public static List<Row> SortRowsToDateAndLessonNumber(List<Row> rows)
    {
        string[] dateFormats = { "dd.MM.yyyy", "dd.MM.yy" };

        return rows
            .OrderBy(row => DateTime.TryParseExact(
                row.Date,
                dateFormats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var date
            ) ? date : DateTime.MaxValue)
            .ThenBy(row => int.TryParse(row.LessonNumber, out int num) ? num : int.MaxValue)
            .ToList();
    }

    static void SaveResultsToExcel(string outputPath)
    {
        try
        {
            var outputFile = new FileInfo(outputPath);

            using (var package = new ExcelPackage(outputFile))
            {
                var worksheet = package.Workbook.Worksheets.Add("Результаты");

                // Заголовки
                string[] headers = Result[0].Split(';');
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headers[i].Trim('"');
                }

                // Данные
                for (int i = 1; i < Result.Count; i++)
                {
                    string[] cells = Result[i].Split(';');
                    for (int j = 0; j < cells.Length; j++)
                    {
                        worksheet.Cells[i + 1, j + 1].Value = cells[j].Trim('"');
                    }
                }

                // Авто-ширина колонок
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                package.Save();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при сохранении результатов: {ex.Message}");
        }
    }

    static string ExtractDateFromFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            return "Дата не определена";

        // Регулярное выражение для поиска даты в начале строки
        var match = System.Text.RegularExpressions.Regex.Match(
            fileName,
            @"^(?<word>\w+\s+)?(?<date>\d{2}\.\d{2}\.(\d{4}|\d{2}))");

        if (match.Success)
        {
            string dateStr = match.Groups["date"].Value;
            string[] formats = { "dd.MM.yyyy", "dd.MM.yy" };

            if (DateTime.TryParseExact(
                dateStr,
                formats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out DateTime date))
            {
                return date.ToString("dd.MM.yyyy");
            }
        }

        return "Дата не определена";
    }
}