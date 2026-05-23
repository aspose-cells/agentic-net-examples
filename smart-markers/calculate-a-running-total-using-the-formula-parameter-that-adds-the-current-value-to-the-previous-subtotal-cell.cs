using System;
using Aspose.Cells;

namespace RunningTotalExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column A (values to be summed)
            // A1: Header, A2..A6: numeric values
            cells["A1"].PutValue("Value");
            double[] values = { 10, 20, 30, 40, 50 };
            for (int i = 0; i < values.Length; i++)
            {
                // Row index is i+2 because rows are 1‑based in the address string
                cells[i + 1, 0].PutValue(values[i]); // Column 0 = A
            }

            // Add header for running total in column B
            cells["B1"].PutValue("Running Total");

            // Set formulas for running total:
            // B2 = A2
            // B3 = B2 + A3
            // B4 = B3 + A4, etc.
            // First row uses only the current value
            cells["B2"].Formula = "=A2";

            // Subsequent rows reference the previous total cell
            for (int row = 3; row <= values.Length + 1; row++)
            {
                // Example for row 3: "=B2+A3"
                string prevTotalCell = $"B{row - 1}";
                string currentValueCell = $"A{row}";
                cells[row - 1, 1].Formula = $"={prevTotalCell}+{currentValueCell}";
            }

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("RunningTotal.xlsx");

            // Optional: display results in console
            Console.WriteLine("Running total calculated:");
            for (int i = 0; i < values.Length; i++)
            {
                Console.WriteLine($"Row {i + 2}: Value={cells[i + 1, 0].Value}, Running Total={cells[i + 1, 1].Value}");
            }
        }
    }
}