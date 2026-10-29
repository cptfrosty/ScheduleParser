namespace ScheduleParser;

public sealed record ScheduleEntry(
    DateOnly Date,
    string LessonNumber,
    string Group,
    string Subgroup,
    string Subject,
    string Teacher,
    string Classroom,
    string Territory,
    string SourceFile,
    string SheetName,
    int ExcelRow);

public sealed record ScheduleSheet(string Name, IReadOnlyList<string[]> Rows)
{
    public string Cell(int row, int column)
    {
        if (row < 0 || row >= Rows.Count || column < 0 || column >= Rows[row].Length)
            return string.Empty;
        return Rows[row][column];
    }
}
