using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyWithStyle
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string destPath = "destination.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Define the source range to be copied (e.g., A1:D5)
                AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:D5");

                // Create a new workbook that will receive the copied range
                Workbook destWorkbook = new Workbook(); // creates a new empty workbook
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Define the destination range with the same size as the source range
                AsposeRange destRange = destSheet.Cells.CreateRange("A1:D5");

                // Copy the source range to the destination range (values, formulas, formatting, drawings)
                sourceRange.Copy(destRange);

                // Ensure the destination directory exists
                string destDir = Path.GetDirectoryName(Path.GetFullPath(destPath));
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                // Save the destination workbook
                destWorkbook.Save(destPath);
                Console.WriteLine($"Range copied successfully. Destination saved to {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}