// Title: Aspose.Cells for .NET: Create a Named Range, Apply a SUM Formula, and Verify the Result in C#
// Description: This example shows how to build a new workbook, fill cells A1‑A3 with numbers, define a named range called MyRange, set cell B1 to =SUM(MyRange), calculate all formulas, retrieve the computed value, and programmatically confirm that the sum equals 60 using Aspose.Cells for .NET.
// Keywords: Aspose.Cells named range C# | SUM formula Aspose.Cells | calculate workbook formulas .NET | verify cell value programmatically | Aspose.Cells example C# | unit test formula result Aspose
// Common Searches: how to add a named range in Aspose.Cells C# | use SUM with a named range Aspose.Cells | calculate and read formula result Aspose.Cells .NET | validate workbook calculations programmatically | Aspose.Cells example for SUM(MyRange)
// Developer Intent: Create a named range, use it in a SUM formula, compute the workbook, and assert that the result matches the expected total.
// Use Cases: Automated financial totals that reference dynamic data blocks | Unit‑testing spreadsheet calculations in CI pipelines | Reusable reporting templates with named ranges and built‑in validation
// AI Prompts: Provide C# code that defines a named range, assigns =SUM(MyRange) to a cell, runs workbook.CalculateFormula, and asserts the result with Aspose.Cells. | Show how to retrieve the calculated value of a cell after formula evaluation in Aspose.Cells for .NET. | Explain how to write a unit test that verifies a SUM formula using a named range returns the correct total.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example shows how to build a new workbook, fill cells A1‑A3 with numbers, define a named range called MyRange, set cell B1 to =SUM(MyRange), calculate all formulas, retrieve the computed value, and programmatically confirm that the sum equals 60 using Aspose.Cells for .NET.
    public class NamedRangeSumDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate some data that will be summed
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["A3"].PutValue(30);

                // Create a named range that refers to A1:A3
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                Name myRange = workbook.Worksheets.Names[nameIndex];
                myRange.RefersTo = "=Sheet1!$A$1:$A$3";

                // Set a formula that uses the named range
                cells["B1"].Formula = "=SUM(MyRange)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Retrieve the calculated result
                object result = cells["B1"].Value;

                // Verify the result programmatically (expected sum = 60)
                double expected = 10 + 20 + 30;
                if (Convert.ToDouble(result) == expected)
                {
                    Console.WriteLine($"Success: SUM(MyRange) = {result}");
                }
                else
                {
                    Console.WriteLine($"Failure: Expected {expected} but got {result}");
                }

                // (Optional) Save the workbook if you want to inspect it
                // workbook.Save("NamedRangeSumDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            NamedRangeSumDemo.Run();
        }
    }
}
