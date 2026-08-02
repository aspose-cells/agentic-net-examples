using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate sample Gregorian dates
        dataSheet.Cells["A1"].PutValue(new DateTime(2023, 1, 15));
        dataSheet.Cells["A2"].PutValue(new DateTime(2023, 2, 20));
        dataSheet.Cells["B1"].PutValue(new DateTime(2023, 3, 25));
        dataSheet.Cells["B2"].PutValue("Not a date"); // non‑date cell

        // Set workbook region to Japan (Japanese locale)
        workbook.Settings.Region = CountryCode.Japan;

        // Create a style that formats dates in Japanese pattern
        Style japaneseStyle = workbook.CreateStyle();
        japaneseStyle.Custom = "[$-ja-JP]yyyy年m月d日";

        // List to hold report entries: cell address, original Gregorian value, Japanese formatted string
        List<(string Address, DateTime Gregorian, string Japanese)> report = new List<(string, DateTime, string)>();

        // Scan used cells for DateTime values, apply Japanese style, and record information
        int maxRow = dataSheet.Cells.MaxDataRow;
        int maxCol = dataSheet.Cells.MaxDataColumn;
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = dataSheet.Cells[row, col];
                if (cell.Type == CellValueType.IsDateTime)
                {
                    DateTime originalDate = cell.DateTimeValue;          // Gregorian value
                    cell.SetStyle(japaneseStyle);                        // Apply Japanese formatting
                    string japaneseFormatted = cell.StringValue;         // Formatted string after style
                    report.Add((cell.Name, originalDate, japaneseFormatted));
                }
            }
        }

        // Create a new worksheet to hold the report
        Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
        reportSheet.Name = "Report";

        // Write header
        reportSheet.Cells["A1"].PutValue("Cell");
        reportSheet.Cells["B1"].PutValue("Gregorian");
        reportSheet.Cells["C1"].PutValue("Japanese");

        // Populate report rows
        for (int i = 0; i < report.Count; i++)
        {
            int rowIndex = i + 1; // offset for header
            reportSheet.Cells[rowIndex, 0].PutValue(report[i].Address);
            reportSheet.Cells[rowIndex, 1].PutValue(report[i].Gregorian);
            reportSheet.Cells[rowIndex, 2].PutValue(report[i].Japanese);
        }

        // Save the workbook with the converted dates and the report
        workbook.Save("JapaneseDateReport.xlsx");
    }
}