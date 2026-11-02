// Title: Evaluate an Excel formula string with custom CalculationOptions using Worksheet.CalculateFormula in Aspose.Cells for .NET
// Common Searches: how to calculate a formula string with CalculationOptions in Aspose.Cells C# | using Worksheet.CalculateFormula to evaluate SUM without inserting into a worksheet | Aspose.Cells recursive calculation option for in‑memory formula evaluation | programmatically compute Excel formulas in .NET without creating cells | evaluate Excel expression with custom options using Aspose.Cells API
// Tags: Worksheet.CalculateFormula with custom CalculationOptions | in‑memory Excel formula evaluation Aspose.Cells | recursive calculation mode Aspose.Cells | C# evaluate SUM function without worksheet cell | Aspose.Cells programmatic formula computation .NET

using Aspose.Cells;
using System;

// The example creates an in‑memory Workbook, configures a CalculationOptions object (enabling recursive calculation), defines a formula string (e.g., SUM(1,2,3)), and calls Worksheet.CalculateFormula(formula, calcOptions) to evaluate the expression without adding it to the sheet, then outputs the result.
class Program
{
    static void Main()
    {
        // Create a new workbook (in-memory) and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Create custom calculation options
        CalculationOptions calcOptions = new CalculationOptions();
        // Example option: enable recursive calculation
        calcOptions.Recursive = true;

        // Formula string to evaluate
        string formula = "SUM(1, 2, 3)";

        // Evaluate the formula without adding it to the worksheet
        object result = sheet.CalculateFormula(formula, calcOptions);

        // Display the result
        Console.WriteLine("Result: " + result);
    }
}
