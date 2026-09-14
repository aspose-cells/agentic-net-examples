// Title: Evaluate a SUM(IF) array formula with Aspose.Cells Worksheet.CalculateFormula in C# without full worksheet recalculation
// AI Prompts: Write C# code that loads an Excel workbook and uses Worksheet.CalculateFormula to compute the result of a SUM(IF) array expression. | Show how to call Aspose.Cells to evaluate a single complex formula while preventing a complete worksheet recalculation.
// Common Searches: Aspose.Cells calculate single formula without recalculating whole sheet C# | C# evaluate SUM(IF) array formula using Worksheet.CalculateFormula | How to get result of complex Excel formula in Aspose.Cells without full recalculation | CalculateFormula example for array formulas in .NET | Avoid full worksheet recalculation when evaluating formula with Aspose.Cells
// Tags: Worksheet.CalculateFormula evaluate array formula | Aspose.Cells single formula evaluation C# | prevent full worksheet recalculation Aspose.Cells | SUM IF array expression .NET | Excel formula computation without full recalculation | complex formula evaluation Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an existing workbook, accesses the first worksheet, defines a complex SUM(IF) array formula as a string, evaluates it using Worksheet.CalculateFormula (which returns the result as an object), and prints the result to the console, all without triggering a full worksheet recalculation.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define a complex formula (example: array formula)
        string complexFormula = "SUM(IF(A1:A10>5, A1:A10, 0))";

        // Evaluate the formula without recalculating the entire worksheet
        // CalculateFormula returns the result as an object
        object result = sheet.CalculateFormula(complexFormula);

        // Output the result
        Console.WriteLine("Result of the formula: " + result);
    }
}
