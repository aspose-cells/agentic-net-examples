// Title: Load an Excel workbook from a UNC network share, adjust compatibility mode, and save it locally with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load workbook from UNC path and save to local folder | How to set Excel compatibility mode when using Aspose.Cells in .NET | Saving an Aspose.Cells workbook to a specific directory after verifying file existence | Create output directory programmatically before saving an Aspose.Cells workbook | Handle FileNotFoundException for network Excel files with Aspose.Cells
// Tags: load workbook from network share using Aspose.Cells | set Excel2007 compatibility mode Aspose.Cells | save workbook as Xlsx to local directory C# | ensure output folder exists before workbook save | catch missing network file exception Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for an Excel file on a UNC network share, loads it into an Aspose.Cells Workbook, optionally sets Excel2007 compatibility mode, creates the destination folder if it does not exist, and then saves the workbook as an .xlsx file to a local distribution path, with basic error handling.
    class Program
    {
        static void Main()
        {
            // Path to the workbook on the network share
            string networkPath = @"\\Server\Share\Folder\source.xlsx";

            // Local path where the modified workbook will be saved
            string localPath = @"C:\Distribution\output.xlsx";

            try
            {
                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(networkPath))
                {
                    Console.WriteLine($"Source file not found: {networkPath}");
                    return;
                }

                // Load the workbook from the network location
                Workbook workbook = new Workbook(networkPath);

                // Optional: set compatibility mode if the API supports it
                // workbook.Settings.Compatibility = CompatibilityMode.Excel2007;

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(localPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook locally for distribution
                workbook.Save(localPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook successfully saved to: {localPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
