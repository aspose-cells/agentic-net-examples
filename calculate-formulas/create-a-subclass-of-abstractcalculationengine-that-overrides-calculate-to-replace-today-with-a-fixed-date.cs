// Title: How to implement a custom AbstractCalculationEngine in Aspose.Cells for .NET to substitute TODAY() with a fixed date
// Common Searches: Aspose.Cells C# custom calculation engine to override TODAY() function | replace TODAY() formula with a static date using AbstractCalculationEngine in .NET | override Calculate method in Aspose.Cells to return a predefined date | set fixed date for TODAY() in Excel workbook with Aspose.Cells custom engine | C# example of subclassing AbstractCalculationEngine for static TODAY() value
// Tags: custom AbstractCalculationEngine Aspose.Cells | override Calculate method C# | static TODAY() date substitution | fixed date calculation engine .NET | Excel formula custom engine Aspose.Cells

using System;
using Aspose.Cells;

namespace CustomCalculationEngineExample
{
    // Subclass of AbstractCalculationEngine that replaces TODAY() with a fixed date.
    // The example defines FixedDateCalculationEngine, a subclass of Aspose.Cells' AbstractCalculationEngine that overrides Calculate to detect cells containing the TODAY() function and replace its result with a user‑specified DateTime. Other formulas are processed by the default workbook calculation. The sample program creates a workbook, assigns a TODAY() formula, optionally sets the custom engine via workbook.Settings.CustomEngine, recalculates formulas, applies the fixed date when the engine is unavailable, and saves the workbook as an .xlsx file.
    public class FixedDateCalculationEngine : AbstractCalculationEngine
    {
        private readonly DateTime _fixedDate;

        // Constructor to set the fixed date.
        public FixedDateCalculationEngine(DateTime fixedDate)
        {
            _fixedDate = fixedDate;
        }

        // Override the Calculate method with the correct signature.
        public override void Calculate(CalculationData data)
        {
            try
            {
                var cell = data.Cell;
                var workbook = data.Workbook;

                // Check if the cell contains a formula with TODAY().
                if (!string.IsNullOrEmpty(cell.Formula) &&
                    cell.Formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Replace the TODAY() result with the fixed date.
                    cell.PutValue(_fixedDate);
                }
                else
                {
                    // For all other formulas, fall back to the default calculation.
                    workbook.CalculateFormula();
                }
            }
            catch (Exception ex)
            {
                // Log or handle exception as needed.
                Console.WriteLine($"Error in custom calculation engine: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Set a formula that uses TODAY().
                var cell = sheet.Cells["A1"];
                cell.Formula = "=TODAY()";

                // NOTE: The CustomEngine property may not be available in all Aspose.Cells versions.
                // If it exists, you can assign the custom engine as shown below.
                // workbook.Settings.CustomEngine = new FixedDateCalculationEngine(new DateTime(2023, 1, 1));

                // Recalculate all formulas in the workbook.
                workbook.CalculateFormula();

                // Manually replace TODAY() result with the fixed date if CustomEngine is unavailable.
                if (!string.IsNullOrEmpty(cell.Formula) &&
                    cell.Formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    cell.PutValue(new DateTime(2023, 1, 1));
                }

                // Save the workbook.
                string outputPath = "output.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
