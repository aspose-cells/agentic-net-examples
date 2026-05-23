using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class RunningTotalExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Set header titles for the data and running total columns
        cells["A1"].PutValue("Value");
        cells["B1"].PutValue("RunningTotal");

        // Create a table (ListObject) that initially covers only the header row (A1:B1)
        // The last two parameters are the end row and end column (both 0 because only header now)
        int tableIndex = sheet.ListObjects.Add(0, 0, 0, 1, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Sample data to be imported row by row
        int[] data = { 10, 20, 15, 30, 25 };

        // Import each value and set a running‑total formula for the same row
        for (int i = 0; i < data.Length; i++)
        {
            // Row offset inside the table: header is 0, first data row is 1, etc.
            int rowOffset = i + 1;          // 1‑based offset for data rows
            int valueColumnOffset = 0;      // Column A inside the table
            int totalColumnOffset = 1;      // Column B inside the table

            // Put the numeric value into the "Value" column
            table.PutCellValue(rowOffset, valueColumnOffset, data[i]);

            // Build a running‑total formula that sums from the first data cell (A2)
            // up to the current row's "Value" cell.
            // Excel rows are 1‑based, so we add 1 to the zero‑based row index.
            int excelRowNumber = rowOffset + 1; // +1 because Excel rows start at 1
            string formula = $"=SUM(A$2:A{excelRowNumber})";

            // Assign the formula to the "RunningTotal" column of the current row
            table.PutCellFormula(rowOffset, totalColumnOffset, formula);
        }

        // Calculate all formulas so that the running totals are materialized
        workbook.CalculateFormula();

        // Save the workbook (adjust the path as needed)
        workbook.Save("RunningTotalExample.xlsx");
    }
}