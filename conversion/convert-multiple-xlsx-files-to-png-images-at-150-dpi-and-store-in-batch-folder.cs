using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace BatchXlsxToPng
{
    public class Converter
    {
        // Converts all .xlsx files in the source directory to PNG images at 150 DPI.
        // Each worksheet is saved as a separate PNG file in the destination directory.
        public static void Run(string sourceFolder, string destinationFolder)
        {
            // Ensure destination folder exists
            Directory.CreateDirectory(destinationFolder);

            // Set the global DPI to 150 for rendering
            CellsHelper.DPI = 150;

            // Get all .xlsx files in the source folder
            string[] files = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                // Load the workbook (lifecycle rule)
                Workbook workbook = new Workbook(filePath);

                // Iterate through each worksheet in the workbook
                for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
                {
                    Worksheet sheet = workbook.Worksheets[sheetIndex];

                    // Configure image rendering options
                    ImageOrPrintOptions options = new ImageOrPrintOptions
                    {
                        ImageType = ImageType.Png,          // Output format
                        OnePagePerSheet = true              // Render whole sheet on one page
                    };

                    // Create a SheetRender instance (lifecycle rule)
                    SheetRender sheetRender = new SheetRender(sheet, options);

                    // Render the first (and only) page to a PNG file
                    // Use the overload ToImage(int pageIndex, string fileName) (rule)
                    string outputFileName = Path.Combine(
                        destinationFolder,
                        $"{Path.GetFileNameWithoutExtension(filePath)}_Sheet{sheetIndex + 1}.png");

                    sheetRender.ToImage(0, outputFileName);

                    // Clean up resources for the current sheet
                    sheetRender.Dispose();
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string sourceDir = @"C:\InputXlsx";
            string destDir   = @"C:\OutputPng";

            Converter.Run(sourceDir, destDir);
        }
    }
}