// Title: C# utility to extract embedded SVG pictures from an Excel workbook with Aspose.Cells and save them to a folder
// Description: A ready‑to‑run C# example that loads an Excel workbook, scans every worksheet for picture shapes, and uses Aspose.Cells' ImageOrPrintOptions (ImageType.Svg) to export each picture as an SVG file. The utility creates the target directory, names files by sheet and shape index, and handles errors gracefully.
// Keywords: Aspose.Cells SVG extraction | C# export Excel picture to SVG | extract embedded images from workbook | save Excel shapes as SVG files | Aspose.Cells ImageOrPrintOptions SVG | .NET Excel SVG utility | batch convert Excel pictures to SVG | GitHub Aspose.Cells SVG extractor
// Common Searches: how to export Excel pictures as SVG using Aspose.Cells | C# code to extract all SVG images from a workbook | save embedded Excel shapes to SVG folder .NET | Aspose.Cells extract SVG from multiple sheets | sample program to convert Excel pictures to SVG
// Developer Intent: Programmatically locate every picture shape in an Excel workbook and write each one as an individual SVG file to a user‑specified directory.
// Use Cases: Create web‑ready vector assets from Excel charts stored as picture shapes. | Automate archival of vector graphics embedded in financial or engineering reports. | Generate SVG resources from template workbooks for UI rendering in .NET applications. | Batch‑process a library of workbooks to produce SVG assets for documentation pipelines.
// AI Prompts: Write a method that scans an Aspose.Cells Workbook for picture shapes and saves each as an SVG file with a name that includes the worksheet name and shape index. | Extend the SvgExtractor to log the original shape name, dimensions, and output path to a CSV file for audit purposes. | Create a PowerShell script that calls the SvgExtractor for every .xlsx file in a directory and stores the resulting SVGs in matching subfolders. | Modify the utility to also export chart objects as SVG using Aspose.Cells' ToImage method.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsUtilities
{
    // A ready‑to‑run C# example that loads an Excel workbook, scans every worksheet for picture shapes, and uses Aspose.Cells' ImageOrPrintOptions (ImageType.Svg) to export each picture as an SVG file. The utility creates the target directory, names files by sheet and shape index, and handles errors gracefully.
    public static class SvgExtractor
    {
        /// <param name="workbookPath">Full path to the source workbook.</param>
        /// <param name="outputFolder">Folder where extracted SVG files will be saved.</param>
        public static void Extract(string workbookPath, string outputFolder)
        {
            try
            {
                // Verify that the workbook file exists.
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Ensure the output directory exists.
                Directory.CreateDirectory(outputFolder);

                // Load the workbook from the provided file path.
                using (Workbook workbook = new Workbook(workbookPath))
                {
                    // Iterate through each worksheet in the workbook.
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        ShapeCollection shapes = sheet.Shapes;

                        // Examine each shape.
                        for (int i = 0; i < shapes.Count; i++)
                        {
                            // Only process picture shapes.
                            if (shapes[i] is Picture picture)
                            {
                                // Export the picture as an SVG file.
                                string fileName = $"Sheet{sheet.Index}_Shape{i}.svg";
                                string outputPath = Path.Combine(outputFolder, fileName);

                                // Configure rendering options for SVG output.
                                ImageOrPrintOptions options = new ImageOrPrintOptions
                                {
                                    ImageType = ImageType.Svg
                                };

                                try
                                {
                                    picture.ToImage(outputPath, options);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Failed to export shape {i} on sheet {sheet.Name}: {ex.Message}");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during SVG extraction: {ex.Message}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string sourceWorkbook = @"C:\Data\SampleWorkbook.xlsx";
            string svgOutputFolder = @"C:\Data\ExtractedSvgs";

            SvgExtractor.Extract(sourceWorkbook, svgOutputFolder);

            Console.WriteLine("SVG extraction completed.");
        }
    }
}
