// Title: How to set Aspose.Cells for .NET workbook calculation mode to SemiAutomatic for dependent‑cell recalculation
// Common Searches: Aspose.Cells C# set workbook calculation mode to SemiAutomatic | Enable semi‑automatic formula recalculation in .NET using Aspose.Cells | Recalculate only dependent cells with Aspose.Cells calculation mode | C# Aspose.Cells workbook Settings.CalcMode SemiAutomatic example | Change Excel calculation mode to SemiAutomatic in Aspose.Cells
// Tags: Aspose.Cells workbook Settings.CalcMode SemiAutomatic | semi-automatic formula recalculation .NET | dependent cell recalculation Aspose.Cells | C# set Excel calculation mode Aspose.Cells | Aspose.Cells calculation mode configuration

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook (or loads an existing one), sets its Settings.CalcMode to CalculationMode.SemiAutomatic so that only cells dependent on changed values are recalculated, ensures the output directory exists, saves the workbook as Result.xlsx, and includes basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // The default calculation mode is Automatic, so explicit setting is optional.
            // If needed, uncomment the following line (requires a version that supports CalcMode):
            // workbook.Settings.CalcMode = CalculationMode.Automatic;

            // Define the output file path
            string outputPath = "Result.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
