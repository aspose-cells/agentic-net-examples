using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTotalExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with a header and numeric values
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(120);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(80);
            cells["A4"].PutValue("C");
            cells["B4"].PutValue(150);

            // Define the range of the table (including header and data rows)
            int startRow = 0;      // Row 0 -> A1
            int startColumn = 0;   // Column 0 -> A
            int endRow = 3;        // Row 3 -> A4/B4 (last data row)
            int endColumn = 1;     // Column 1 -> B

            // Add the ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Enable the totals row; Aspose.Cells will add an extra row after the data rows
            table.ShowTotals = true;

            // Calculate the sum of the "Amount" column (column index 1) manually
            double sum = 0;
            // Data rows are from row offset 1 (first data row) to (table.EndRow - table.StartRow - 1) (excluding totals row)
            int dataRowCount = table.EndRow - table.StartRow; // includes totals row, so subtract 1 for data rows
            for (int i = 1; i < dataRowCount; i++) // start at 1 to skip header row
            {
                // Retrieve the cell value from the worksheet using absolute coordinates
                object val = cells[table.StartRow + i, table.StartColumn + 1].Value;
                if (val is double d) sum += d;
                else if (val is int n) sum += n;
            }

            // Determine the offset for the totals row (last row of the table)
            int totalsRowOffset = table.EndRow - table.StartRow; // zero‑based offset; points to the totals row
            int amountColumnOffset = 1; // second column in the table

            // Insert the calculated total value into the totals row using PutCellValue
            table.PutCellValue(totalsRowOffset, amountColumnOffset, sum);

            // Optionally, set a label for the totals row in the first column
            table.PutCellValue(totalsRowOffset, 0, "Grand Total", true);

            // Save the workbook
            workbook.Save("TableWithCalculatedTotal.xlsx", SaveFormat.Xlsx);
        }
    }
}