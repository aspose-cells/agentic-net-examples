// Title: C# – Dynamically Freeze Rows to the Last Data Row with Worksheet.Cells.MaxDataRow in Aspose.Cells
// Description: Demonstrates how to create a workbook, populate it with data, determine the last populated row using Worksheet.Cells.MaxDataRow (zero‑based), and apply FreezePanes to lock all rows above that point before saving the file.
// Keywords: Aspose.Cells dynamic freeze rows | Worksheet.Cells.MaxDataRow C# | FreezePanes based on data range | .NET Excel freeze panes | auto freeze header rows Aspose
// Common Searches: Aspose.Cells freeze rows up to last data row | How to use MaxDataRow with FreezePanes in C# | Dynamic freeze panes Aspose.Cells .NET | Set freeze panes after loading CSV in Aspose.Cells | Auto‑freeze header rows in Excel using Aspose
// Developer Intent: Find the final row containing data and lock all preceding rows in the worksheet automatically.
// Use Cases: Lock header rows after importing an unknown number of records from a database. | Keep the top rows visible in reports generated from CSV files of varying length. | Create printable Excel sheets where the title and column headings stay static regardless of data size.
// AI Prompts: Write C# code that reads data into a worksheet, uses Worksheet.Cells.MaxDataRow to locate the last row, freezes rows up to that point, and saves the workbook. | Explain MaxDataRow vs. MaxDataColumn and show how to apply FreezePanes for both rows and columns dynamically. | Provide an example that loads external data, writes it to an Aspose.Cells worksheet, then applies a dynamic FreezePanes call based on the detected data range.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicFreezeDemo
{
    // Demonstrates how to create a workbook, populate it with data, determine the last populated row using Worksheet.Cells.MaxDataRow (zero‑based), and apply FreezePanes to lock all rows above that point before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (header + several rows)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            for (int i = 2; i <= 10; i++)
            {
                sheet.Cells[i - 1, 0].PutValue(i - 1);                     // ID
                sheet.Cells[i - 1, 1].PutValue($"Item {i - 1}");          // Name
                sheet.Cells[i - 1, 2].PutValue(10 * (i - 1));            // Score
            }

            // Determine the last row that contains data
            int maxDataRow = sheet.Cells.MaxDataRow; // zero‑based index

            // If there is data, freeze all rows up to the last data row
            if (maxDataRow >= 0)
            {
                // Freeze panes just below the last data row.
                // row parameter is the index of the cell where the split occurs,
                // so we use maxDataRow + 1 to place the split after the data.
                // freezedRows specifies how many rows to keep visible in the top pane.
                sheet.FreezePanes(maxDataRow + 1, 0, maxDataRow + 1, 0);
            }

            // Save the workbook
            workbook.Save("DynamicFreezeRows.xlsx");
        }
    }
}
