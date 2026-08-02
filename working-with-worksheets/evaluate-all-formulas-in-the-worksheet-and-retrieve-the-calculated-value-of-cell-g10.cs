// Title: Evaluate all worksheet formulas and read cell G10 value with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, sets numeric values and formulas, runs Worksheet.CalculateFormula with default CalculationOptions to recalculate the entire sheet, then accesses the Value property of G10 to obtain the computed result (130).
// Keywords: Aspose.Cells C# calculate formulas | Worksheet.CalculateFormula example | retrieve cell value after calculation | evaluate all formulas Aspose.Cells | read G10 result .NET
// Common Searches: Aspose.Cells evaluate all formulas C# | how to get calculated value of a cell in Aspose.Cells | Worksheet.CalculateFormula usage example | read result of G10 after formula evaluation | C# Aspose.Cells recalculate worksheet
// Developer Intent: Recalculate every formula in a worksheet and obtain the numeric result of cell G10.
// Use Cases: Generate a report where dependent totals must be refreshed after programmatic data changes. | Extract a specific metric (e.g., G10) for further processing after batch formula evaluation. | Validate that complex formulas produce expected outcomes in automated tests.
// AI Prompts: Show C# code that uses Aspose.Cells to recalculate all formulas and return the value of cell G10. | Demonstrate Worksheet.CalculateFormula with CalculationOptions and how to read the resulting cell value. | Explain handling of different data types (number, string, date) when retrieving a calculated cell value in Aspose.Cells.

using System;
using Aspose.Cells;

namespace EvaluateFormulasExample
{
    // Creates a workbook, sets numeric values and formulas, runs Worksheet.CalculateFormula with default CalculationOptions to recalculate the entire sheet, then accesses the Value property of G10 to obtain the computed result (130).
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creation rule

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Example data: populate some cells with values and formulas
            // ------------------------------------------------------------
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["B1"].Formula = "=SUM(A1:A2)";   // B1 = 15
            sheet.Cells["C1"].Formula = "=B1*2";        // C1 = 30
            sheet.Cells["G10"].Formula = "=C1+100";    // G10 = 130

            // ------------------------------------------------------------
            // Calculate all formulas in the worksheet
            // ------------------------------------------------------------
            // Using the worksheet-level CalculateFormula method with default options
            // This evaluates every formula in the sheet.
            sheet.CalculateFormula(new CalculationOptions(), true); // rule: CalculateFormula(CalculationOptions, bool)

            // ------------------------------------------------------------
            // Retrieve the calculated value of cell G10
            // ------------------------------------------------------------
            Cell targetCell = sheet.Cells["G10"];
            object calculatedValue = targetCell.Value; // Value holds the result after calculation

            // Output the result
            Console.WriteLine($"Calculated value of G10: {calculatedValue}");
        }
    }
}
