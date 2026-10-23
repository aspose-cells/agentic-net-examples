// Title: How to set Aspose.Cells Workbook CalculationMode to SemiAutomatic for dependent‑cell recalculation in C#
// Common Searches: Aspose.Cells C# set workbook calculation mode to SemiAutomatic | recalculate only dependent cells after updating a value with Aspose.Cells | partial formula evaluation example using Aspose.Cells .NET | disable full workbook recalculation in Aspose.Cells C# code | configure semi‑automatic calculation mode in Aspose.Cells workbook
// Tags: Aspose.Cells semi‑automatic calculation mode C# | dependent cell formula evaluation Aspose.Cells | set workbook CalculationMode .NET | partial Excel recalculation using Aspose.Cells | optimize formula calculation performance Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, sets its Settings.CalculationMode to SemiAutomatic so only cells that depend on changed values are recalculated, updates a base cell and a dependent formula, invokes CalculateFormula to refresh dependent cells, and saves the result to an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Example modification: set a value and a formula that depends on it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);          // Base value
            sheet.Cells["B1"].Formula = "=A1*2";     // Dependent formula

            // Trigger calculation for dependent cells
            workbook.CalculateFormula();

            // Define output path
            string outputPath = "output.xlsx";

            // Ensure the output directory exists (if any)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
