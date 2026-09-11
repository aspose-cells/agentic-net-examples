// Title: Add and style a ListObject table (A1:D10) in an existing or new XLSX workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that opens "input.xlsx" if it exists (or creates a new workbook), inserts a ListObject covering A1:D10 with headers, sets its DisplayName to "MyTable", enables the header row, applies TableStyleMedium2, and saves the result as "output.xlsx". | Write a C# snippet using Aspose.Cells to create a 10‑row by 4‑column Excel table, assign a custom display name, show the header row, apply the built‑in TableStyleMedium2, and then save the workbook. | Provide Aspose.Cells for .NET code that adds a ListObject to the first worksheet, configures it with a header row, custom name, and TableStyleMedium2, handling both existing and new workbook scenarios.
// Common Searches: aspnet c# add ListObject to worksheet with Aspose.Cells and apply built‑in table style | how to create a styled Excel table with headers using Aspose.Cells .NET | load existing workbook or create new one then insert table in Aspose.Cells C# example | set display name for Aspose.Cells ListObject and enable header row | apply TableStyleMedium2 to Excel table using Aspose.Cells C#
// Tags: add ListObject table Aspose.Cells C# | apply TableStyleMedium2 Aspose.Cells | set ListObject display name Excel | create styled Excel table Aspose.Cells | load or create workbook Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The program checks for an existing "input.xlsx" file, loads it or creates a new workbook, adds a ListObject covering cells A1:D10 with headers, assigns the display name "MyTable", shows the header row, applies the built‑in TableStyleMedium2 style, and saves the workbook as "output.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // default workbook with one worksheet
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for the table (example: A1:D10).
            int firstRow = 0;          // Row 1 (zero‑based)
            int firstColumn = 0;       // Column A
            int totalRows = 10;        // Number of rows in the list
            int totalColumns = 4;      // Number of columns in the list
            bool hasHeaders = true;    // First row contains headers

            // Add a ListObject (table) to the worksheet.
            int listObjectIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                hasHeaders);

            ListObject table = sheet.ListObjects[listObjectIndex];

            // Assign a display name to the table (compatible with all Aspose.Cells versions).
            table.DisplayName = "MyTable";

            // Ensure the header row is displayed.
            table.ShowHeaderRow = true;

            // Apply a built‑in table style.
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Save the workbook with the new table.
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
