using System;
using Aspose.Cells;

namespace EvaluateRangeFormulas
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("Input.xlsx");

            // Access the first worksheet (or specify by name/index)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to evaluate: A1:D20
            int startRow = 0;      // Zero‑based index for row 1
            int startColumn = 0;   // Zero‑based index for column A
            int endRow = 19;       // Row 20
            int endColumn = 3;     // Column D

            // Prepare calculation options (default settings)
            CalculationOptions calcOptions = new CalculationOptions();

            // Iterate through each cell in the defined range
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startColumn; col <= endColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // If the cell contains a formula, calculate it individually
                    if (cell.IsFormula)
                    {
                        // Calculate only this cell using the provided rule
                        cell.Calculate(calcOptions);
                    }
                }
            }

            // Save the workbook with the evaluated results (replace with desired output path)
            workbook.Save("Output.xlsx");
        }
    }
}