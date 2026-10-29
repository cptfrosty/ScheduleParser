using System.Globalization;
using System.Text;
using ExcelDataReader;

namespace ScheduleParser;

public interface IWorkbookReader
{
    IReadOnlyList<ScheduleSheet> Read(string path);
}

public sealed class ExcelWorkbookReader : IWorkbookReader
{
    public ExcelWorkbookReader()
    {
        // Required for legacy BIFF/.xls files on .NET 8.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public IReadOnlyList<ScheduleSheet> Read(string path)
    {
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var sheets = new List<ScheduleSheet>();

        do
        {
            var rows = new List<string[]>();
            while (reader.Read())
            {
                var row = new string[reader.FieldCount];
                for (var column = 0; column < reader.FieldCount; column++)
                    row[column] = Convert.ToString(reader.GetValue(column), CultureInfo.InvariantCulture)?.Trim() ?? "";
                rows.Add(row);
            }

            // Only propagate vertical merges. Filling horizontal merges would put
            // a group heading in the subject/teacher columns and corrupt the rows.
            foreach (var merge in reader.MergeCells ?? [])
            {
                if (merge.FromRow == merge.ToRow || merge.FromRow >= rows.Count)
                    continue;

                var value = GetCell(rows, merge.FromRow, merge.FromColumn);
                if (string.IsNullOrWhiteSpace(value))
                    continue;

                for (var r = merge.FromRow; r <= merge.ToRow && r < rows.Count; r++)
                for (var c = merge.FromColumn; c <= merge.ToColumn && c < rows[r].Length; c++)
                    if (string.IsNullOrWhiteSpace(rows[r][c]))
                        rows[r][c] = value;
            }

            sheets.Add(new ScheduleSheet(reader.Name, rows));
        } while (reader.NextResult());

        return sheets;
    }

    private static string GetCell(IReadOnlyList<string[]> rows, int row, int column) =>
        row >= 0 && row < rows.Count && column >= 0 && column < rows[row].Length
            ? rows[row][column]
            : string.Empty;
}
