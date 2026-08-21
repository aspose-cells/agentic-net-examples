// Title: Read the calculated value of cell D5 after Worksheet.CalculateFormula in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, assigns numbers to A5‑C5, sets a SUM formula in D5, runs Worksheet.CalculateFormula with recursive evaluation, and retrieves the computed result via the Cell.Value property.
// Keywords: Aspose.Cells C# | Worksheet.CalculateFormula | read formula result | cell value after calculation | SUM(A5:C5) example | CalculationOptions recursive | retrieve calculated cell | Aspose.Cells .NET example
// Common Searches: Aspose.Cells get value after Worksheet.CalculateFormula | C# read calculated result of a formula cell | How to retrieve D5 value after SUM formula in Aspose.Cells | Worksheet.CalculateFormula recursive flag usage | Aspose.Cells read cell value after calculation
// Developer Intent: Obtain the evaluated result of cell D5 after programmatically calculating worksheet formulas.
// Use Cases: Generate a total row in a financial report, calculate it with Worksheet.CalculateFormula, then export the sum from D5. | Automated unit testing of spreadsheet logic by setting input cells, invoking calculation, and asserting D5's value. | Integrate a computed metric (e.g., total sales) from D5 into downstream services after in‑memory workbook processing.
// AI Prompts: Show C# code that sets values in A5‑C5, adds a SUM formula to D5, runs Worksheet.CalculateFormula with recursive option, and returns the cell's Value. | Explain how CalculationOptions and the recursive flag influence formula evaluation before accessing the result in Aspose.Cells. | Provide a step‑by‑step example of reading a formula cell's calculated value after calling Worksheet.CalculateFormula in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a workbook, assigns numbers to A5‑C5, sets a SUM formula in D5, runs Worksheet.CalculateFormula with recursive evaluation, and retrieves the computed result via the Cell.Value property.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some cells that will be used in the formula
        worksheet.Cells["A5"].PutValue(10);
        worksheet.Cells["B5"].PutValue(20);
        worksheet.Cells["C5"].PutValue(30);

        // Set a formula in D5 that depends on the above cells
        worksheet.Cells["D5"].Formula = "=SUM(A5:C5)";

        // Calculate all formulas in this worksheet (recursive = true)
        worksheet.CalculateFormula(new CalculationOptions(), true);

        // Read the calculated result of cell D5
        object calculatedResult = worksheet.Cells["D5"].Value;

        // Output the result
        Console.WriteLine("Calculated value of D5: " + calculatedResult);
    }
}
