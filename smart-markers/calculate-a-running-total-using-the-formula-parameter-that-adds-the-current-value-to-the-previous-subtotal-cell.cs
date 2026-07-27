using System;
using Aspose.Cells;

namespace RunningTotalExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: values in column B (index 1)
            // Header
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Amount");
            cells["C1"].PutValue("Running Total");

            // Populate some numeric values
            int[] amounts = { 100, 250, 175, 300, 225 };
            for (int i = 0; i < amounts.Length; i++)
            {
                // Item name (optional)
                cells[i + 1, 0].PutValue($"Item {i + 1}");
                // Amount value in column B
                cells[i + 1, 1].PutValue(amounts[i]);
            }

            // Set running total formulas in column C (index 2)
            // First row: running total equals the first amount
            cells[1, 2].Formula = "=B2";

            // Subsequent rows: previous total + current amount
            for (int row = 2; row <= amounts.Length; row++) // Excel rows 3..n+1
            {
                // Excel row numbers are 1‑based, so use row for formula strings
                // Previous total is in column C of the previous row (C{row-1})
                // Current amount is in column B of the current row (B{row})
                cells[row, 2].Formula = $"=C{row - 1}+B{row}";
            }

            // Calculate all formulas so that the running totals are materialized
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("RunningTotal.xlsx");
        }
    }
}