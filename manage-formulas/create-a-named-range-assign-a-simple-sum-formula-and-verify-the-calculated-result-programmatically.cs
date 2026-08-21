// Title: Aspose.Cells for .NET: Define a Named Range, Apply a SUM Formula, and Verify the Result Programmatically
// Description: C# sample that creates a workbook, fills A1‑A3 with numbers, defines a workbook‑level named range "MyRange", inserts a SUM formula referencing the range into B1, forces calculation with workbook.CalculateFormula, reads the result, and checks that it equals the expected total of 60.
// Keywords: Aspose.Cells | .NET | C# named range | SUM formula | formula calculation | programmatic verification | workbook.CalculateFormula | Excel automation | named range example | cell value validation
// Common Searches: Aspose.Cells create named range C# | How to use SUM with a named range in Aspose.Cells | Validate formula result after CalculateFormula Aspose.Cells | C# example for workbook.CalculateFormula | Aspose.Cells programmatic Excel calculations
// Developer Intent: Define a named range, apply a SUM formula that references it, calculate the workbook, and confirm the computed value in code.
// Use Cases: Generate financial reports where totals are derived from a reusable named range of expense cells. | Implement data integrity checks by comparing summed values of a named range against predefined thresholds. | Create dynamic dashboards that rely on named ranges whose aggregates are validated before visualization.
// AI Prompts: Write C# code with Aspose.Cells to create a named range for cells C1:C5 and set an AVERAGE formula that uses the range. | Show how to handle an empty named range so that a SUM formula returns zero without throwing an exception in Aspose.Cells. | Demonstrate updating the reference of an existing named range and recalculating dependent formulas using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeDemo
{
    // C# sample that creates a workbook, fills A1‑A3 with numbers, defines a workbook‑level named range "MyRange", inserts a SUM formula referencing the range into B1, forces calculation with workbook.CalculateFormula, reads the result, and checks that it equals the expected total of 60.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (default name is "Sheet1")
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some numeric data that will be summed
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // ----- Create a named range -----
            // Add a new name to the workbook's name collection
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            // Set the range that the name refers to (absolute A1:A3 on Sheet1)
            workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$A$3";

            // ----- Assign a SUM formula that uses the named range -----
            // Place the formula in cell B1
            cells["B1"].Formula = "=SUM(MyRange)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the calculated result from B1
            object result = cells["B1"].Value;

            // Verify the result programmatically (expected sum = 10 + 20 + 30 = 60)
            double expected = 60;
            double actual = Convert.ToDouble(result);

            Console.WriteLine($"Calculated SUM result in B1: {actual}");
            Console.WriteLine($"Verification {(Math.Abs(actual - expected) < 1e-9 ? "passed" : "failed")}.");

            // Optionally save the workbook to inspect the result manually
            workbook.Save("NamedRangeSumDemo.xlsx");
        }
    }
}
