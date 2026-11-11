// Title: Remove all pictures from an Excel workbook after loading with Aspose.Cells for .NET to boost performance
// Common Searches: C# how to delete all images from an Excel file using Aspose.Cells | Aspose.Cells remove pictures from worksheets to speed up large workbook processing | strip pictures from a loaded workbook in .NET with Aspose.Cells | optimize loading of large Excel files by discarding pictures using Aspose.Cells C# | programmatically clear worksheet pictures in Aspose.Cells after opening workbook
// Tags: remove pictures Aspose.Cells C# | strip images from Excel workbook .NET | optimize large workbook loading Aspose.Cells | worksheet picture removal API Aspose.Cells | discard Excel images during processing C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPictureFilter
{
    // The example checks for the source Excel file, loads it with Aspose.Cells, iterates each worksheet, removes every picture by traversing the Pictures collection in reverse, saves the workbook to a new file without images, and includes basic exception handling to improve performance on large spreadsheets.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook.
            string inputPath = "LargeWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook (pictures will be loaded initially).
                Workbook workbook = new Workbook(inputPath);

                // Remove all pictures from each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Collect picture indices because removing modifies the collection.
                    int pictureCount = sheet.Pictures.Count;
                    for (int i = pictureCount - 1; i >= 0; i--)
                    {
                        sheet.Pictures.RemoveAt(i);
                    }
                }

                // Save the workbook after removing pictures.
                string outputPath = "LargeWorkbook_NoPictures.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved without pictures to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
