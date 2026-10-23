// Title: Recalculate only changed cells with Aspose.Cells CalculationOptions (Recursive = false) in C#
// Common Searches: Aspose.Cells calculate only modified cells C# | How to use CalculationOptions Recursive false to limit formula recalculation | Selective formula recalculation after updating cell values with Aspose.Cells .NET | Trigger dependent formula update without full workbook recalculation in C# | CalculateFormula for changed cells only using Aspose.Cells API
// Tags: Aspose.Cells CalculationOptions Recursive false | selective formula recalculation .NET | recalculate dependent cells Aspose.Cells | partial workbook calculation C# | update formulas after cell change Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates a workbook, modifies cell A1, sets a dependent formula in B1, configures CalculationOptions with Recursive = false so that CalculateFormula recomputes only the cells that changed, and then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the input file is missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a workbook with one empty worksheet
            }

            // Example modifications – only these cells have changed
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);               // changed cell
            sheet.Cells["B1"].Formula = "=A1*2";          // dependent formula

            // Recalculate only the cells that changed since the last calculation
            CalculationOptions calcOptions = new CalculationOptions
            {
                // When false, Aspose.Cells recalculates only dependent cells.
                Recursive = false
            };
            workbook.CalculateFormula(calcOptions);

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
