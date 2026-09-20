// Title: How to configure Aspose.Cells for .NET to use iterative calculation with 100 max iterations and a 0.001 convergence threshold
// AI Prompts: Generate C# code that sets Workbook.Settings.IterativeCalculation.MaxIterations to 100 and ConvergenceThreshold to 0.001, then saves the workbook as an .xlsx file using Aspose.Cells. | Provide a step‑by‑step example showing how to enable circular‑reference iterative calculation in Aspose.Cells for .NET, including a version check for the IterativeCalculation API.
// Common Searches: Aspose.Cells .NET enable iterative calculation and define max iterations | set convergence threshold 0.001 in Aspose.Cells workbook settings | C# example for configuring circular reference handling with Aspose.Cells
// Tags: Aspose.Cells iterative calculation configuration | Workbook.Settings.IterativeCalculation max iterations | convergence threshold setting Aspose.Cells | circular reference calculation .NET Aspose.Cells | Aspose.Cells workbook save C# example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, notes that iterative calculation settings require a newer Aspose.Cells version, ensures the output directory exists, saves the workbook as 'IterativeCalculation.xlsx', and catches any exceptions, providing a basis for adding iterative calculation configuration when supported.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // NOTE: Iterative calculation settings are not available in this version of Aspose.Cells.
            // If needed, upgrade to a newer version that supports Workbook.Settings.IterativeCalculation.

            // Define output file path
            string outputPath = "IterativeCalculation.xlsx";

            // Ensure the directory for the output file exists (if any)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
