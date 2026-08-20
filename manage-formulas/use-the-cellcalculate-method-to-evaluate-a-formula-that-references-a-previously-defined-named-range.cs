// Title: Calculate a SUM over a named range with Cell.Calculate in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills A1:A3 with numbers, defines the named range "MyRange", assigns the formula "=SUM(MyRange)" to B1, evaluates it using Cell.Calculate with CalculationOptions, outputs the result, and saves the file.
// Keywords: Aspose.Cells | Cell.Calculate | named range | SUM formula | C# | .NET | CalculationOptions | evaluate formula programmatically | Excel automation
// Common Searches: Aspose.Cells calculate formula with named range | Cell.Calculate C# example | How to evaluate SUM(MyRange) using Aspose.Cells | Retrieve calculated value without opening Excel | Aspose.Cells named range calculation
// Developer Intent: Programmatically evaluate a formula that references a previously defined named range using Cell.Calculate.
// Use Cases: Generate total or average values from a dynamic data set before persisting the workbook. | Validate spreadsheet formulas after updating cells that belong to a named range. | Produce intermediate calculation results for reports where formulas depend on named ranges. | Unit‑test spreadsheet logic by asserting calculated values in code.
// AI Prompts: Provide C# code that defines a named range, sets a SUM formula, and uses Cell.Calculate to obtain the result. | Explain how to configure CalculationOptions for custom precision when evaluating named‑range formulas in Aspose.Cells. | Show error‑handling patterns for Cell.Calculate when the referenced named range is missing or invalid.

using System;
using Aspose.Cells;

// Creates a workbook, fills A1:A3 with numbers, defines the named range "MyRange", assigns the formula "=SUM(MyRange)" to B1, evaluates it using Cell.Calculate with CalculationOptions, outputs the result, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data in cells A1:A3
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["A3"].PutValue(30);

        // Define a named range "MyRange" that refers to A1:A3
        int nameIndex = workbook.Worksheets.Names.Add("MyRange");
        workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$A$3";

        // Set a formula in cell B1 that references the named range
        Cell formulaCell = worksheet.Cells["B1"];
        formulaCell.Formula = "=SUM(MyRange)";

        // Evaluate the formula using Cell.Calculate
        formulaCell.Calculate(new CalculationOptions());

        // Output the calculated result
        Console.WriteLine("SUM(MyRange) = " + formulaCell.Value);

        // Save the workbook (saving rule)
        workbook.Save("NamedRangeCalculation.xlsx");
    }
}
