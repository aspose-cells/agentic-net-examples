// Title: Calculate a SUM formula referencing a named range with Cell.Calculate in Aspose.Cells for .NET
// Description: Creates a workbook, defines a named range (MyRange) over A1:A3, sets =SUM(MyRange) in B1, and uses Cell.Calculate with default CalculationOptions to obtain the result before optionally saving the file.
// Keywords: Aspose.Cells | Cell.Calculate | named range | SUM formula | C# | .NET | CalculationOptions | programmatic calculation | workbook automation
// Common Searches: Cell.Calculate named range Aspose.Cells | Aspose.Cells calculate SUM of named range C# | evaluate formula that uses a named range in Aspose.Cells | recalculate cells after defining named ranges programmatically | Aspose.Cells calculate cell value without opening Excel
// Developer Intent: Recalculate a cell containing a formula that references a previously defined named range using Aspose.Cells.
// Use Cases: Generate summary totals for dynamic data blocks defined by named ranges. | Validate intermediate results in a workbook before persisting it. | Create on‑the‑fly calculations when populating reports programmatically.
// AI Prompts: Show C# code that defines a named range, assigns a SUM formula, and retrieves the calculated value with Cell.Calculate. | Explain how CalculationOptions affect precision when evaluating a named‑range formula in Aspose.Cells. | Provide steps to recalculate a cell after updating values inside its referenced named range.

using System;
using Aspose.Cells;

// Creates a workbook, defines a named range (MyRange) over A1:A3, sets =SUM(MyRange) in B1, and uses Cell.Calculate with default CalculationOptions to obtain the result before optionally saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some sample data in column A
        cells["A1"].PutValue(5);
        cells["A2"].PutValue(10);
        cells["A3"].PutValue(15);

        // Define a named range "MyRange" that refers to A1:A3
        int nameIndex = workbook.Worksheets.Names.Add("MyRange");
        Name namedRange = workbook.Worksheets.Names[nameIndex];
        namedRange.RefersTo = "=Sheet1!$A$1:$A$3";

        // Set a formula in cell B1 that uses the named range
        Cell formulaCell = cells["B1"];
        formulaCell.Formula = "=SUM(MyRange)";

        // Evaluate the formula using Cell.Calculate with default options
        formulaCell.Calculate(new CalculationOptions());

        // Output the calculated result
        Console.WriteLine("SUM(MyRange) = " + formulaCell.Value);

        // Save the workbook (optional)
        workbook.Save("NamedRangeCalculation.xlsx");
    }
}
