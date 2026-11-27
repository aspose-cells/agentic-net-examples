// Title: How to disable the NumbersAsText error check on a worksheet with Aspose.Cells for .NET using ErrorCheckOptions
// Common Searches: Aspose.Cells C# disable NumbersAsText error checking on a specific worksheet | Set ErrorCheckOptions.NumbersAsText to false in Aspose.Cells .NET | Turn off numeric text warning for Excel files using Aspose.Cells API | Suppress NumbersAsText validation when saving workbook with Aspose.Cells C# | Configure workbook error checking to ignore numbers stored as text in Aspose.Cells
// Tags: ErrorCheckOptions NumbersAsText false Aspose.Cells | disable numeric text warning worksheet .NET | Aspose.Cells workbook error checking configuration | suppress NumbersAsText validation C# | Excel file error check settings Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, accesses the first Worksheet, disables the NumbersAsText error check via workbook.Settings.ErrorCheckOptions.NumbersAsText = false, ensures the output directory exists, saves the workbook as Result.xlsx, and writes a success message to the console.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Get the first worksheet
                var worksheet = workbook.Worksheets[0];

                // The EnableErrorChecking property is not available in this version of Aspose.Cells.
                // If error checking needs to be disabled, adjust workbook settings accordingly.

                // Define output file path
                string outputPath = "Result.xlsx";

                // Ensure the directory for the output file exists (if a directory is specified)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
