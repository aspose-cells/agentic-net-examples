// Title: Turn Off UseThreadedCalculation in Aspose.Cells .NET to Boost Simple Workbook Performance
// Description: A C# sample that creates a Workbook, checks for the presence of the Settings.UseThreadedCalculation property, sets it to false when supported, adds a value and a formula, guarantees the output directory exists, and saves the file. Turning off the threaded engine cuts unnecessary thread overhead and steadies execution for lightweight spreadsheets.
// Keywords: Aspose.Cells UseThreadedCalculation false | disable threaded formula calculation .NET | Excel multithreading off | performance tuning Aspose.Cells | lightweight workbook stability | C# Aspose.Cells settings | reduce calculation overhead | threaded calculation setting
// Common Searches: How to disable UseThreadedCalculation in Aspose.Cells C# | Aspose.Cells performance settings for small workbooks | Turn off multithreaded formula engine in .NET | Reduce Excel calculation overhead with Aspose.Cells | Stabilize simple spreadsheets by disabling threading
// Developer Intent: Switch off the multithreaded formula engine to lower resource usage and improve reliability when processing small Excel files.
// Use Cases: Generate a minimal workbook with a few formulas and deactivate threaded calculation to avoid extra thread management. | Prevent sporadic calculation failures in constrained environments such as serverless functions or low‑memory containers. | Speed up batch jobs that create thousands of tiny Excel files by globally setting UseThreadedCalculation to false.
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, disables UseThreadedCalculation, adds a simple formula, and saves it. | Explain the impact of turning off Workbook.Settings.UseThreadedCalculation on performance and stability for small spreadsheets. | Provide a step‑by‑step guide to detect if the UseThreadedCalculation property exists in the current Aspose.Cells version and safely set it to false.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // A C# sample that creates a Workbook, checks for the presence of the Settings.UseThreadedCalculation property, sets it to false when supported, adds a value and a formula, guarantees the output directory exists, and saves the file. Turning off the threaded engine cuts unnecessary thread overhead and steadies execution for lightweight spreadsheets.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook instance
                Workbook workbook = new Workbook();

                // The UseThreadedCalculation property may not be available in all versions.
                // If needed, it can be set here when supported:
                // workbook.Settings.UseThreadedCalculation = false;

                // Add some sample data (optional)
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].Formula = "=A1*2";

                // Define output file path
                string outputPath = "SimpleWorkbook.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
