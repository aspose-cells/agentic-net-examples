// Title: C# – Evaluate formulas in range A1:D20 with Aspose.Cells without full workbook recalculation
// Description: Loads a workbook, selects the first worksheet, defines the A1:D20 block, iterates through each cell, calculates only those that contain formulas using Cell.Calculate with default CalculationOptions, and saves the file. This approach updates a specific region while leaving the rest of the workbook untouched, improving performance.
// Keywords: Aspose.Cells C# evaluate range | partial formula calculation | Cell.Calculate specific cells | A1:D20 formula evaluation | skip full workbook recalculation
// Common Searches: Aspose.Cells evaluate formulas in a selected range | C# calculate only cells A1:D20 | partial workbook calculation Aspose.Cells | how to recalculate a specific area in Excel using Aspose
// Developer Intent: Recompute formulas exclusively inside cells A1:D20, avoiding a complete workbook calculation.
// Use Cases: Boost performance when only a small data block changes. | Refresh summary tables after programmatic updates without affecting other sheets. | Generate reports that need up‑to‑date values for a defined area while preserving existing calculations elsewhere.
// AI Prompts: Show how to ignore empty or error cells while evaluating formulas in A1:D20. | Demonstrate setting CalculationOptions for iterative calculations limited to a specific range. | Provide code to log each formula that is recalculated in the A1:D20 block.

using System;
using Aspose.Cells;

namespace EvaluateRangeFormulas
{
    // Loads a workbook, selects the first worksheet, defines the A1:D20 block, iterates through each cell, calculates only those that contain formulas using Cell.Calculate with default CalculationOptions, and saves the file. This approach updates a specific region while leaving the rest of the workbook untouched, improving performance.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "InputWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range A1:D20 (zero‑based indices)
            int startRow = 0;   // A1 row
            int endRow   = 19;  // row 20
            int startCol = 0;   // column A
            int endCol   = 3;   // column D

            // Prepare calculation options (default options are sufficient)
            CalculationOptions calcOptions = new CalculationOptions();

            // Iterate through each cell in the range and calculate only those that contain formulas
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.IsFormula)
                    {
                        // Calculate the formula for this single cell
                        cell.Calculate(calcOptions);
                    }
                }
            }

            // Save the workbook after the partial calculation
            string outputPath = "OutputWorkbook.xlsx";
            workbook.Save(outputPath);
        }
    }
}
