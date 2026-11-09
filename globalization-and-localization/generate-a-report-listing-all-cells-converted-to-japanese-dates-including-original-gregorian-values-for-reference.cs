// Title: Generate a Japanese calendar date report with original Gregorian values using Aspose.Cells for .NET (C#)
// Common Searches: how to create a Japanese era date report from Excel using Aspose.Cells in C# | C# Aspose.Cells example to list Gregorian dates and their Japanese calendar equivalents | iterate through all worksheets and export DateTime cells to Japanese calendar format with Aspose.Cells | Aspose.Cells generate report of cells containing dates and convert to JapaneseCalendar | save Excel workbook with a new sheet showing Japanese era dates alongside original dates in C#
// Tags: Aspose.Cells JapaneseCalendar date conversion | Excel DateTime cell enumeration Aspose.Cells C# | generate date conversion report worksheet Aspose.Cells | format dates with ja-JP culture in Aspose.Cells | auto-fit columns after writing report Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// The sample loads an existing workbook, adds a worksheet named "JapaneseDateReport", and writes headers for sheet name, cell address, Gregorian value, and Japanese date. It configures a ja-JP CultureInfo with a JapaneseCalendar, iterates every worksheet (excluding the report sheet) and its used range, identifies cells of type DateTime, converts each Gregorian date to a Japanese era string, records both values in the report sheet, auto‑fits the columns, and saves the workbook as "output.xlsx".
class JapaneseDateReportGenerator
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create a new worksheet for the report
        Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
        reportSheet.Name = "JapaneseDateReport";

        // Write header row
        Cells reportCells = reportSheet.Cells;
        reportCells["A1"].PutValue("Sheet Name");
        reportCells["B1"].PutValue("Cell Address");
        reportCells["C1"].PutValue("Gregorian Value");
        reportCells["D1"].PutValue("Japanese Date");

        // Prepare Japanese culture with JapaneseCalendar
        CultureInfo jaCulture = new CultureInfo("ja-JP");
        jaCulture.DateTimeFormat.Calendar = new JapaneseCalendar();

        int reportRow = 1; // zero‑based index; row 1 is the second row (after header)

        // Iterate through all worksheets and cells
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Skip the report sheet itself
            if (sheet.Name == reportSheet.Name) continue;

            Cells cells = sheet.Cells;
            // Use the MaxDataRow and MaxDataColumn to limit the iteration to used range
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];
                    // Check if the cell contains a DateTime value
                    if (cell.Type == CellValueType.IsDateTime)
                    {
                        DateTime gregorianDate = cell.DateTimeValue;

                        // Convert to Japanese calendar string
                        string japaneseDate = gregorianDate.ToString("gg y年M月d日", jaCulture);

                        // Write the information to the report sheet
                        reportCells[reportRow, 0].PutValue(sheet.Name);                     // Sheet Name
                        reportCells[reportRow, 1].PutValue(cell.Name);                     // Cell Address (e.g., "B2")
                        reportCells[reportRow, 2].PutValue(gregorianDate);                 // Gregorian Value
                        reportCells[reportRow, 3].PutValue(japaneseDate);                  // Japanese Date

                        reportRow++;
                    }
                }
            }
        }

        // Auto‑fit columns for better readability
        reportSheet.AutoFitColumns();

        // Save the workbook with the report
        workbook.Save("output.xlsx");
    }
}
