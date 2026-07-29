// Title: C# – Verify and display formulas in the second worksheet with Aspose.Cells
// Description: Load an Excel file, access the worksheet at index 1, calculate all formulas, and print each formula cell’s address and value using Aspose.Cells for .NET. Optionally save the workbook after verification.
// Keywords: Aspose.Cells C# | second worksheet formula evaluation | calculate formulas Aspose.Cells | iterate cells with formulas | read cell value after calculation | save workbook after verification | Excel formula verification .NET
// Common Searches: Aspose.Cells get second worksheet and evaluate formulas | C# iterate cells in a specific sheet to read formula results | How to calculate all formulas before reading values in Aspose.Cells | Save workbook after formula verification using Aspose.Cells | Aspose.Cells example for formula verification in .NET
// Developer Intent: Access the second worksheet, compute its formulas once, and output the results.
// Use Cases: Confirm that every formula in a target sheet calculates correctly before distribution. | Generate a log of formula cells with their evaluated values for auditing or debugging. | Update and persist the workbook after ensuring formula accuracy.
// AI Prompts: Write C# code with Aspose.Cells to open a workbook, select the second worksheet, calculate formulas once, and print each formula cell’s address and value. | Provide an optimized version that calculates formulas a single time and writes the results to a CSV file. | Explain how to handle workbooks that may lack a second worksheet when using Aspose.Cells, including error handling best practices.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaVerification
{
    // Load an Excel file, access the worksheet at index 1, calculate all formulas, and print each formula cell’s address and value using Aspose.Cells for .NET. Optionally save the workbook after verification.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the second worksheet (zero‑based index 1)
            Worksheet secondWorksheet = workbook.Worksheets[1];

            // Prepare for formula verification:
            // Iterate through all cells in the second worksheet,
            // identify those containing formulas, calculate them,
            // and output the resulting values.
            Cells cells = secondWorksheet.Cells;
            foreach (Cell cell in cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Calculate all formulas in the workbook (required before reading values)
                    workbook.CalculateFormula();

                    // Display the cell address and its calculated value
                    Console.WriteLine($"Cell {cell.Name} = {cell.Value}");
                }
            }

            // Save the workbook after verification (optional)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}
