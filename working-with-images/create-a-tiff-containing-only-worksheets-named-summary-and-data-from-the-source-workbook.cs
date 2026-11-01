// Title: Create a multi‑page TIFF from only the 'Summary' and 'Data' worksheets using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export selected Excel sheets to a multi‑page TIFF | How to generate a TIFF file containing only specific worksheets with Aspose.Cells | Save only certain worksheets as TIFF images in .NET | Aspose.Cells copy particular sheets before saving as TIFF in C# | Multi‑page TIFF export of selected Excel tabs using Aspose.Cells
// Tags: selected worksheets to TIFF conversion Aspose.Cells | copy specific sheets to new workbook C# | multi-page TIFF generation from Excel Aspose.Cells | filter Excel tabs before image export .NET | save workbook as TIFF with chosen sheets

using Aspose.Cells;
using System;
using System.IO;

// The example loads a source workbook, extracts the 'Summary' and 'Data' worksheets into a new workbook, removes the default sheet, and saves the result as a multi‑page TIFF file. It ensures the output directory exists and handles missing source files or sheets gracefully.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            var sourcePath = "source.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            var sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook that will contain only the required sheets
            var newWorkbook = new Workbook();

            // Remove the default empty worksheet that Aspose.Cells creates
            newWorkbook.Worksheets.Clear();

            // Names of worksheets to keep
            string[] sheetsToCopy = { "Summary", "Data" };

            // Copy each required worksheet from the source workbook to the new workbook
            foreach (var sheetName in sheetsToCopy)
            {
                var srcSheet = sourceWorkbook.Worksheets[sheetName];
                if (srcSheet != null)
                {
                    // Add a new empty worksheet to the target workbook
                    int newIndex = newWorkbook.Worksheets.Add();
                    var newSheet = newWorkbook.Worksheets[newIndex];

                    // Set the name to match the source sheet
                    newSheet.Name = srcSheet.Name;

                    // Copy the contents from the source sheet to the new sheet
                    newSheet.Copy(srcSheet);
                }
                else
                {
                    Console.WriteLine($"Worksheet \"{sheetName}\" not found in source workbook.");
                }
            }

            // Path for the output TIFF file
            var outputPath = "output.tiff";

            // Ensure the directory for the output file exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the new workbook as a multi‑page TIFF (each worksheet becomes a page)
            newWorkbook.Save(outputPath, SaveFormat.Tiff);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
