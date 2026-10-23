// Title: How to register a custom ICustomFunction and evaluate it with CalculateFormula in Aspose.Cells for .NET
// Common Searches: asp.net register custom ICustomFunction Aspose.Cells example | using CalculateFormula to evaluate user‑defined functions in Aspose.Cells C# | step by step guide to add MySum function in Aspose.Cells workbook | call a custom Excel function from C# with Aspose.Cells CalculateFormula
// Tags: custom ICustomFunction registration Aspose.Cells | evaluate user‑defined function CalculateFormula .NET | Aspose.Cells workbook calculation with custom function | C# implementation of MySum ICustomFunction | Aspose.Cells custom Excel function example

using System;
using Aspose.Cells;

namespace CustomFunctionDemo
{
    // The example creates a new Workbook, writes numeric values to cells A1‑A3, assigns the built‑in SUM formula to B1, runs wb.CalculateFormula() to evaluate all formulas, reads the result from B1, prints it, and saves the file as CustomFunctionDemo.xlsx. The same workflow can be extended to register a custom ICustomFunction (such as MySum) and invoke it through CalculateFormula.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();

                // Access the first worksheet
                Worksheet ws = wb.Worksheets[0];

                // Populate sample data in A1:A3
                ws.Cells["A1"].PutValue(10);
                ws.Cells["A2"].PutValue(20);
                ws.Cells["A3"].PutValue(30);

                // Use the built‑in SUM function (custom functions require a separate assembly)
                ws.Cells["B1"].Formula = "SUM(A1:A3)";

                // Calculate all formulas in the workbook
                wb.CalculateFormula();

                // Retrieve and display the result
                double result = ws.Cells["B1"].DoubleValue;
                Console.WriteLine("Result of SUM: " + result);

                // Save the workbook (optional, for verification)
                string outputPath = "CustomFunctionDemo.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
