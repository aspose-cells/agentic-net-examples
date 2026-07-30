// Title: Add a Grand Total to an Aspose.Cells ListObject Table with PutCellValue (C# .NET)
// Description: Creates a workbook, converts a range to a ListObject, enables the totals row, computes the sum of the "Amount" column, and writes the total and a label into the totals row using the correct row offset via ListObject.PutCellValue.
// Keywords: Aspose.Cells | ListObject | PutCellValue | totals row | C# | .NET | table sum | grand total | Excel automation
// Common Searches: Aspose.Cells add total to table row | ListObject PutCellValue offset example | C# calculate column sum in Aspose.Cells table | how to enable totals row Aspose.Cells | write grand total to Aspose.Cells ListObject
// Developer Intent: Insert a calculated total and optional label into the totals row of a ListObject table using PutCellValue with the proper offset.
// Use Cases: Generate sales reports that automatically display the sum of the Amount column at the bottom of the table. | Create inventory sheets where the quantity column is totaled and labeled in the table's totals row. | Build financial statements that dynamically compute and show column totals without manual editing.
// AI Prompts: Show C# code that sums a numeric column in an Aspose.Cells ListObject and writes the result to the totals row using PutCellValue. | Explain how to calculate the row offset for a ListObject totals row and insert both a label and the computed total. | Provide a step‑by‑step example of enabling a totals row, iterating over data rows, and populating the totals row with PutCellValue.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTotalExample
{
    // Creates a workbook, converts a range to a ListObject, enables the totals row, computes the sum of the "Amount" column, and writes the total and a label into the totals row using the correct row offset via ListObject.PutCellValue.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (header + 4 data rows)
            // Column A: Item, Column B: Amount
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Date");
            sheet.Cells["B5"].PutValue(50);

            // Define the range that will become the table (including header)
            int startRow = 0;      // Row 0 -> A1
            int startColumn = 0;   // Column 0 -> A
            int endRow = 5;        // Row 5 -> A6 (last data row)
            int endColumn = 1;     // Column 1 -> B

            // Add the ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Enable the totals row for the table
            table.ShowTotals = true;

            // Calculate the total of the "Amount" column manually
            double total = 0;
            // Data rows start after the header (row offset 1) and end before the totals row.
            // At this point the totals row has just been added, so EndRow includes it.
            // Therefore, iterate up to EndRow - 1.
            for (int r = table.StartRow + 1; r < table.EndRow; r++)
            {
                object val = sheet.Cells[r, 1].Value; // Column B (index 1)
                if (val is double d) total += d;
                else if (val is int i) total += i;
            }

            // Determine the row offset for the totals row within the table
            int totalsRowOffset = table.EndRow - table.StartRow; // zero‑based offset

            // Insert the calculated total value into the totals row, column "Amount" (index 1)
            table.PutCellValue(totalsRowOffset, 1, total);

            // Optionally, set a label in the first column of the totals row
            table.PutCellValue(totalsRowOffset, 0, "Grand Total", true);

            // Save the workbook
            workbook.Save("TableWithCalculatedTotal.xlsx", SaveFormat.Xlsx);
        }
    }
}
