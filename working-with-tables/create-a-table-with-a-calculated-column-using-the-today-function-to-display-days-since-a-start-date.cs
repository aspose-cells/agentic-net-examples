// Title: Add a calculated column that shows days since a start date using TODAY in an Excel table with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create table with calculated column using TODAY function | how to compute days elapsed from a start date in an Excel table via Aspose.Cells | C# Aspose.Cells structured reference formula for days since date | generate Excel workbook with ListObject and dynamic days‑since column in .NET | using TODAY() in Aspose.Cells table formula to calculate age of dates
// Tags: Aspose.Cells calculated column TODAY formula | C# ListObject table structured reference | Excel table days‑since column Aspose.Cells | Aspose.Cells auto‑fit columns workbook | Aspose.Cells .NET table style TableStyleMedium9

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, adds a ListObject table with "StartDate" and "DaysSince" headers, fills start dates, defines a calculated column "DaysSince" using the formula "=TODAY()-[StartDate]", applies a table style, auto‑fits columns, and saves the file as TableWithDaysSince.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Set headers for the table.
            sheet.Cells["A1"].PutValue("StartDate");
            sheet.Cells["B1"].PutValue("DaysSince");

            // Populate some start dates.
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["A3"].PutValue(new DateTime(2023, 3, 15));
            sheet.Cells["A4"].PutValue(new DateTime(2023, 6, 30));
            sheet.Cells["A5"].PutValue(new DateTime(2023, 12, 1));

            // Define the range that will become a table (ListObject).
            // The range includes headers and data rows.
            int firstRow = 0;   // zero‑based index for row 1
            int firstCol = 0;   // zero‑based index for column A
            int totalRows = 5;  // header + 4 data rows
            int totalCols = 2;  // StartDate and DaysSince

            // Add the table to the worksheet (hasHeaders = true).
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstCol,
                firstRow + totalRows - 1,
                firstCol + totalCols - 1,
                true);

            ListObject table = sheet.ListObjects[tableIndex];
            // Header row is already shown because hasHeaders = true.
            // Totals row is not required for this example.
            table.TableStyleType = TableStyleType.TableStyleMedium9; // optional styling

            // Set the name of the calculated column.
            table.ListColumns[1].Name = "DaysSince";

            // Apply a calculated formula using the TODAY function.
            // The formula uses a structured reference to the StartDate column.
            table.ListColumns[1].Formula = "=TODAY()-[StartDate]";

            // Auto‑fit columns for better visibility.
            sheet.AutoFitColumns();

            // Determine output path.
            string outputPath = "TableWithDaysSince.xlsx";

            // Save the workbook to a file.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
