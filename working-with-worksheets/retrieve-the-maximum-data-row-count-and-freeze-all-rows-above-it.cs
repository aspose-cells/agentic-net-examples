// Title: Freeze rows up to the last populated row using Aspose.Cells for .NET (C#)
// Description: Shows how to retrieve the highest row containing data with Cells.MaxDataRow and apply Worksheet.FreezePanes to lock all rows from the top of the sheet through that row, then save the workbook as an XLSX file.
// Keywords: Aspose.Cells | Cells.MaxDataRow | Worksheet.FreezePanes | freeze top rows | C# Excel automation | Aspose.Cells .NET | freeze rows programmatically | last data row | Excel freeze panes C# | dynamic freeze rows
// Common Searches: Aspose.Cells freeze rows up to last data row | How to use Cells.MaxDataRow in C# | Freeze top rows in Excel with Aspose.Cells .NET | Worksheet.FreezePanes example C# | Get maximum data row Aspose.Cells | Dynamic freeze panes Aspose.Cells | C# code to freeze header rows in Excel
// Developer Intent: Identify the final row that contains data and programmatically freeze every row above (including) it in an Excel worksheet.
// Use Cases: Generating reports where header and summary rows must stay visible while scrolling | Creating templates that adapt to varying data lengths and keep the top section static | Automating Excel exports that require a frozen pane based on the amount of data inserted | Building dashboards where the first N rows need to remain on screen regardless of scroll position
// AI Prompts: Provide C# code using Aspose.Cells to detect the last non‑empty row and freeze all rows above it. | Create a reusable function that accepts a Worksheet object and applies FreezePanes from row 0 to Cells.MaxDataRow. | Explain the relationship between Cells.MaxDataRow and Worksheet.FreezePanes for dynamic row freezing in Excel files.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to retrieve the highest row containing data with Cells.MaxDataRow and apply Worksheet.FreezePanes to lock all rows from the top of the sheet through that row, then save the workbook as an XLSX file.
    public class FreezeRowsAboveMaxDataRow
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
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // Sample data – in real scenario the workbook may already contain data
            // -------------------------------------------------
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Item 1");
            cells["A3"].PutValue("Item 2");
            cells["A4"].PutValue("Item 3");
            cells["A5"].PutValue("Item 4");
            // -------------------------------------------------

            // Retrieve the maximum data row index (zero‑based)
            int maxDataRow = cells.MaxDataRow;   // -1 if no data
            Console.WriteLine("Maximum data row index: " + maxDataRow);

            if (maxDataRow >= 0)
            {
                // Freeze all rows above (and including) the max data row.
                // FreezePanes(row, column, freezedRows, freezedColumns)
                // Set the freeze position to the first row after the data (maxDataRow + 1)
                // and freeze the rows from 0 to maxDataRow.
                int freezeRow = maxDataRow + 1;   // first unfrozen row
                int freezeColumn = 0;             // no column freeze
                int frozenRows = maxDataRow + 1;  // number of rows to freeze
                int frozenColumns = 0;            // no column freeze

                sheet.FreezePanes(freezeRow, freezeColumn, frozenRows, frozenColumns);
                Console.WriteLine($"Frozen top {frozenRows} rows.");
            }
            else
            {
                Console.WriteLine("Worksheet contains no data; nothing to freeze.");
            }

            // Save the workbook
            workbook.Save("FreezeRowsAboveMaxDataRow.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved as FreezeRowsAboveMaxDataRow.xlsx");
        }
    }
}
