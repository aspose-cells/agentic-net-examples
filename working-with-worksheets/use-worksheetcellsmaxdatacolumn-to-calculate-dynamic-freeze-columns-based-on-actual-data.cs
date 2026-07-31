// Title: C# – Dynamically Freeze Columns with Worksheet.Cells.MaxDataColumn in Aspose.Cells
// Description: The sample builds a workbook, fills a few rows, retrieves the index of the right‑most non‑empty column via Worksheet.Cells.MaxDataColumn, and then calls FreezePanes to lock the header row and every column that contains data before saving the file.
// Keywords: Aspose.Cells | C# | Worksheet.Cells.MaxDataColumn | FreezePanes | dynamic column freeze | Excel export automation | zero‑based column index | programmatic freeze panes
// Common Searches: Aspose.Cells freeze columns up to last used column | MaxDataColumn example C# | How to set FreezePanes based on data range in Aspose.Cells | Dynamic freeze panes for Excel export .NET | Get last populated column Aspose.Cells
// Developer Intent: Identify the furthest column containing data and programmatically freeze the top row together with all populated columns in an Aspose.Cells worksheet.
// Use Cases: Exporting database query results where column count varies, while keeping the header row and left‑most fields visible. | Generating reports that require the left side of the sheet to stay in view regardless of how many data columns are added. | Automating spreadsheet creation for dashboards where new metrics may introduce additional columns, and the freeze area must adapt automatically.
// AI Prompts: Write C# code that uses Worksheet.Cells.MaxDataColumn to freeze the first row and every column that has data in an Aspose.Cells workbook. | Show how to safely handle a worksheet with no data before calling FreezePanes with MaxDataColumn. | Compare MaxDataColumn and MaxColumn in Aspose.Cells and demonstrate why MaxDataColumn is preferred for dynamic freeze pane logic.

using System;
using Aspose.Cells;

namespace DynamicFreezeColumnsDemo
{
    // The sample builds a workbook, fills a few rows, retrieves the index of the right‑most non‑empty column via Worksheet.Cells.MaxDataColumn, and then calls FreezePanes to lock the header row and every column that contains data before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data: populate some cells with varying column counts
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["C1"].PutValue("Email");
            worksheet.Cells["D1"].PutValue("Phone");
            worksheet.Cells["E1"].PutValue("Address");

            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("John Doe");
            worksheet.Cells["C2"].PutValue("john@example.com");
            worksheet.Cells["D2"].PutValue("123-456-7890");
            worksheet.Cells["E2"].PutValue("123 Main St");

            // Determine the maximum column index that contains data
            int maxDataColumn = worksheet.Cells.MaxDataColumn; // zero‑based index

            // If there is data, freeze the top row and all columns up to the last data column
            if (maxDataColumn >= 0)
            {
                // Freeze position is the cell just after the frozen area.
                // Row index = 1 (freeze first row), Column index = maxDataColumn + 1 (freeze columns 0..maxDataColumn)
                int freezeRow = 1;                     // first visible row after frozen rows
                int freezeColumn = maxDataColumn + 1;  // first visible column after frozen columns
                int frozenRows = 1;                    // number of rows to freeze (top row)
                int frozenColumns = maxDataColumn + 1; // number of columns to freeze (from column A)

                worksheet.FreezePanes(freezeRow, freezeColumn, frozenRows, frozenColumns);
            }

            // Save the workbook
            workbook.Save("DynamicFreezeColumnsDemo.xlsx");
        }
    }
}
