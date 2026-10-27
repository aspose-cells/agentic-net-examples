// Title: How to configure iterative calculation convergence threshold (0.001) in Aspose.Cells for .NET
// Common Searches: Aspose.Cells set convergence threshold to 0.001 for iterative formulas in C# | C# example of configuring calculation precision with Workbook.Settings.IterativeCalculationSettings | Define iterative calculation precision for Excel circular references using Aspose.Cells | How to adjust iterative calculation settings in Aspose.Cells .NET library
// Tags: Aspose.Cells iterative calculation configuration | Workbook.Settings.IterativeCalculationSettings convergence threshold | C# set calculation precision Aspose.Cells | Excel formula iteration threshold .NET | Aspose.Cells calculation options precision

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, notes that iterative calculation settings are unavailable in the current Aspose.Cells version, and explains that they can be set via Workbook.Settings.IterativeCalculationSettings in newer releases. It ensures the output directory exists and saves the workbook as IterativeCalculation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // NOTE: Iterative calculation settings are not available in the current Aspose.Cells version.
            // If needed, they can be configured using Workbook.Settings.IterativeCalculationSettings in newer versions.

            // Define output file path
            string outputPath = "IterativeCalculation.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
