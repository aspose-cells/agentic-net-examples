// Title: Dynamic Freeze Rows with Worksheet.Cells.MaxDataRow in Aspose.Cells for .NET
// Description: Creates a workbook, populates a worksheet, uses Worksheet.Cells.MaxDataRow to find the last populated row (zero‑based) and applies FreezePanes to lock all rows up to that point, then saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | Worksheet.Cells.MaxDataRow | dynamic freeze rows | freeze panes programmatically | last data row Excel | Excel automation | FreezePanes API | zero‑based row index
// Common Searches: Aspose.Cells freeze rows up to last data row | Worksheet.Cells.MaxDataRow example C# | dynamic freeze panes based on data Aspose | how to set FreezePanes programmatically in .NET | freeze header rows automatically Aspose.Cells
// Developer Intent: Identify the final row containing data and programmatically freeze all rows above it.
// Use Cases: Keep header and summary rows visible in reports whose row count varies. | Automatically apply freeze panes after importing CSV or database data. | Generate Excel templates where the number of rows changes per user input.
// AI Prompts: Generate C# code that uses Aspose.Cells to freeze rows up to the last non‑empty row. | Show how to read a DataTable into a worksheet and apply a dynamic freeze pane using MaxDataRow. | Explain how to extend the logic to also freeze columns based on Worksheet.Cells.MaxDataColumn.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, populates a worksheet, uses Worksheet.Cells.MaxDataRow to find the last populated row (zero‑based) and applies FreezePanes to lock all rows up to that point, then saves the file as an Excel workbook.
    public class DynamicFreezeRowsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (you can replace this with your own data source)
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Item1");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Item2");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("Item3");
            sheet.Cells["B4"].PutValue(300);

            // Determine the last row that contains data (zero‑based index)
            int maxDataRow = sheet.Cells.MaxDataRow;

            // Freeze rows up to the last data row if any data exists
            if (maxDataRow >= 0)
            {
                // Freeze rows: maxDataRow + 1 rows (since row index is zero‑based)
                sheet.FreezePanes(maxDataRow + 1, 0, maxDataRow + 1, 0);
            }

            // Save the workbook
            workbook.Save("DynamicFreezeRowsDemo.xlsx");
        }
    }
}
