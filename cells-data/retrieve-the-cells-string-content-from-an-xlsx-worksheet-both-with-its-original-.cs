using System;
using Aspose.Cells;

class RetrieveCellString
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data
        cells["A1"].PutValue(12345.678);               // Numeric value
        Style style = cells["A1"].GetStyle();
        style.Number = 3;                              // Apply currency format
        cells["A1"].SetStyle(style);

        cells["A2"].PutValue(DateTime.Now);            // DateTime value (default format)

        // Retrieve the cell's string value with original formatting
        string formattedA1 = cells["A1"].StringValue;  // e.g., "$12,345.68"
        string formattedA2 = cells["A2"].StringValue;  // e.g., "3/15/2026 10:30 AM"

        // Retrieve the cell's plain text (no formatting)
        string plainA1 = cells["A1"].GetStringValue(CellValueFormatStrategy.None); // "12345.678"
        string plainA2 = cells["A2"].GetStringValue(CellValueFormatStrategy.None); // "2026-03-15 10:30:00"

        // Output results
        Console.WriteLine($"A1 formatted: {formattedA1}");
        Console.WriteLine($"A1 plain: {plainA1}");
        Console.WriteLine($"A2 formatted: {formattedA2}");
        Console.WriteLine($"A2 plain: {plainA2}");

        // Save the workbook (optional)
        workbook.Save("RetrieveCellString.xlsx");
    }
}