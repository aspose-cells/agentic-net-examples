// Title: Recalculate all formulas in an Excel workbook using Aspose.Cells for .NET (Workbook.CalculateFormula with automatic calculation mode)
// Common Searches: Aspose.Cells calculate all formulas in a workbook C# | How to force full formula recalculation with Aspose.Cells .NET | Set CalculationOptions.CalculationMode to Automatic before CalculateFormula | Reevaluate Excel formulas after loading file using Aspose.Cells | Trigger workbook-wide formula evaluation in C# Aspose.Cells
// Tags: Workbook.CalculateFormula full workbook recalculation | Aspose.Cells automatic calculation mode | set CalculationOptions.CalculationMode C# | recalculate Excel formulas Aspose.Cells | load or create workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file (or creates a new workbook if the file is missing), relies on the default Automatic calculation mode, invokes Workbook.CalculateFormula to recompute every formula in the workbook, and then saves the updated workbook to the specified output path.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Recalculate all formulas in the workbook (default calculation mode is Automatic).
            workbook.CalculateFormula();

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook after recalculation.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
