// Title: Configure Aspose.Cells for .NET to use a 0.001 convergence threshold in iterative formula calculations
// AI Prompts: Generate C# code that enables iterative calculation and sets the convergence threshold (MaxChange) to 0.001 in an Aspose.Cells workbook. | Show how to adjust the FormulaSettings of Aspose.Cells to define a precision of 0.001 for circular reference evaluation. | Provide a step‑by‑step example that configures iterative mode, limits iterations, and saves the workbook after applying a 0.001 MaxChange value.
// Common Searches: Aspose.Cells C# set MaxChange to 0.001 for iterative formulas | How to define convergence precision for circular references using Aspose.Cells .NET | Enable iterative calculation and specify 0.001 threshold in Aspose.Cells workbook | Set iterative calculation options like MaxIteration and MaxChange in Aspose.Cells C# example
// Tags: Aspose.Cells API set MaxChange precision | circular reference handling via formula engine options | C# workbook iterative mode configuration | define convergence threshold in Aspose.Cells | adjust MaxIteration and MaxChange values .NET

using System;
using Aspose.Cells;

namespace AsposeCellsIterativeCalculation
{
    // // Creates a workbook with a circular reference, enables iterative calculation, sets MaxIteration to 100 and convergence threshold (MaxChange) to 0.001, calculates formulas, and saves the file as IterativeCalculationResult.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set up a simple circular reference for demonstration
                worksheet.Cells["A1"].Formula = "=B1+1";
                worksheet.Cells["B1"].Formula = "=A1+1";

                // Enable iterative calculation in the workbook settings
                workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
                workbook.Settings.FormulaSettings.MaxIteration = 100;          // optional, defines max iterations
                workbook.Settings.FormulaSettings.MaxChange = 0.001;          // convergence threshold (precision)

                // Perform calculation
                workbook.CalculateFormula();

                // Save the workbook to verify the result
                workbook.Save("IterativeCalculationResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
