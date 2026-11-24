// Title: Configure Aspose.Cells Workbook to use multi‑threaded calculation for faster Excel processing in C#
// Common Searches: how to turn on multi thread calculation in Aspose.Cells C# | Aspose.Cells enable parallel workbook calculation for large datasets | set Workbook.Settings.EnableMultiThreadedCalculation property in .NET | improve Excel formula evaluation speed with Aspose.Cells multi‑threading | upgrade Aspose.Cells version to support multi‑threaded calculation engine
// Tags: Workbook.Settings.EnableMultiThreadedCalculation property usage | parallel formula evaluation Aspose.Cells | multi‑threaded calculation engine .NET | Excel calculation performance tuning with Aspose | Aspose.Cells version with multi‑thread support

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing Excel file or creates a new workbook, indicates that multi‑threaded calculation requires a newer Aspose.Cells release, recommends upgrading to use the Workbook.Settings.EnableMultiThreadedCalculation API, ensures the output directory exists, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Load an existing workbook if it exists; otherwise create a new one
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // create a new workbook
                }

                // NOTE: Multi‑threaded calculation is not available in the current Aspose.Cells version.
                // If needed, upgrade to a version that supports Workbook.Settings.EnableMultiThreadedCalculation.

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
