// Title: How to Register a Custom AbstractCalculationEngine with Workbook.CalculationEngine in Aspose.Cells for .NET
// Description: Shows how to create a workbook, fill cells, assign a user‑defined AbstractCalculationEngine to the Workbook.CalculationEngine property, and then invoke CalculateFormula to evaluate the formula.
// Keywords: Aspose.Cells | AbstractCalculationEngine | Workbook.CalculationEngine | .NET | custom formula engine | override calculation | user‑defined functions | performance optimization | Excel calculation engine
// Common Searches: set custom calculation engine Aspose.Cells | use AbstractCalculationEngine .NET | replace default formula calculator Aspose | register calculation engine before CalculateFormula | custom function support Aspose.Cells
// Developer Intent: Integrate a bespoke calculation engine into a workbook and run formula evaluation.
// Use Cases: Add proprietary functions not available in the built‑in engine | Boost calculation speed for very large spreadsheets | Log each calculation step for auditing purposes | Enforce specific rounding or numeric precision rules | Switch between multiple engines dynamically at runtime
// AI Prompts: Generate C# code that defines a class inheriting AbstractCalculationEngine and registers it with Workbook.CalculationEngine before calling CalculateFormula. | Show how to implement a custom function inside the overridden Calculate method for Aspose.Cells. | Explain steps to restore the default calculation engine after using a custom one. | Provide an example of logging each cell calculation using a custom engine.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineDemo
{
    // Shows how to create a workbook, fill cells, assign a user‑defined AbstractCalculationEngine to the Workbook.CalculationEngine property, and then invoke CalculateFormula to evaluate the formula.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();

                // Prepare some data and a formula that multiplies the two cells
                Worksheet ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue(5);   // first operand
                ws.Cells["A2"].PutValue(7);   // second operand
                ws.Cells["A3"].Formula = "=A1*A2"; // use built‑in multiplication instead of a custom function

                // Perform calculation
                wb.CalculateFormula();

                // Output the result
                Console.WriteLine("Result of A1*A2: " + ws.Cells["A3"].Value);

                // Save the workbook (optional, just to demonstrate that saving works)
                string outputPath = "CustomEngineResult.xlsx";
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
