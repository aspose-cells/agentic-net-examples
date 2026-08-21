// Title: C# – Evaluate all formulas in an Excel worksheet and retrieve cell G10 value using Aspose.Cells
// Description: Load an Excel file with Aspose.Cells for .NET, run Worksheet.CalculateFormula (including cross‑sheet recursion), and read the computed result of cell G10. The sample prints the value and optionally saves the workbook.
// Keywords: Aspose.Cells formula evaluation | C# calculate all formulas | Worksheet.CalculateFormula example | retrieve Excel cell value G10 | .NET Excel calculation | cross‑worksheet formula recursion | Aspose.Cells sample code
// Common Searches: how to calculate all formulas in Aspose.Cells worksheet | C# Aspose.Cells get value of G10 after calculation | Worksheet.CalculateFormula recursive option example | Aspose.Cells .NET read calculated cell value | evaluate Excel formulas with Aspose.Cells and fetch a specific cell
// Developer Intent: Run a full formula recalculation on a workbook and obtain the resulting value of cell G10.
// Use Cases: Generate a report that requires the final value of a summary cell after all dependent formulas are evaluated. | Validate calculation results in automated tests by comparing the computed G10 value against expected data. | Store the evaluated G10 result in a database or pass it to downstream business logic without manually opening Excel.
// AI Prompts: Show a C# snippet that loads an Excel file, calls Worksheet.CalculateFormula with recursion, and prints the value of cell G10. | Explain how to handle calculation errors in Aspose.Cells and safely cast the G10 result to a double. | Provide a step‑by‑step guide to evaluate all formulas in a workbook and retrieve a specific cell value for use in a .NET application.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaEvaluation
{
    // Load an Excel file with Aspose.Cells for .NET, run Worksheet.CalculateFormula (including cross‑sheet recursion), and read the computed result of cell G10. The sample prints the value and optionally saves the workbook.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains formulas
            string inputPath = "input.xlsx";

            // Create and load the workbook (lifecycle rule)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare calculation options (default options)
            CalculationOptions calcOptions = new CalculationOptions();

            // Calculate all formulas in the worksheet (rule: Worksheet.CalculateFormula)
            // The second parameter 'true' enables recursive calculation across worksheets.
            worksheet.CalculateFormula(calcOptions, true);

            // Retrieve the calculated value of cell G10
            Cell targetCell = worksheet.Cells["G10"];
            object g10Value = targetCell.Value;

            // Output the result
            Console.WriteLine($"Calculated value of G10: {g10Value}");

            // (Optional) Save the workbook after calculation
            // workbook.Save("output.xlsx");
        }
    }
}
