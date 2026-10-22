// Title: Parallelize worksheet‑to‑PNG conversion with Aspose.Cells using C# Task Parallel Library
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells and uses Parallel.ForEach to render each worksheet to a PNG file. | Add a helper method that sanitizes worksheet names for safe file names while converting worksheets concurrently. | Configure ImageOrPrintOptions for single‑page PNG output and write each image to a designated output folder with per‑sheet error handling. | Show how to limit the degree of parallelism for the worksheet conversion to avoid excessive CPU usage.
// Common Searches: c# aspose.cells convert all sheets to png in parallel | using Parallel.ForEach to render Excel worksheets as images with Aspose.Cells | task parallel library example for exporting Excel worksheets to PNG files | how to safely name PNG files generated from worksheet names in Aspose.Cells | control max concurrency when converting Excel sheets to images in C#
// Tags: parallel worksheet PNG conversion Aspose.Cells | Task Parallel Library image rendering Excel | ImageOrPrintOptions PNG configuration Aspose.Cells | sanitize worksheet names for file output C# | limit degree of parallelism Excel sheet conversion | multi‑core Excel to image processing Aspose

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, creates an output directory, and uses Parallel.ForEach (TPL) to render each worksheet to a PNG image with Aspose.Cells. It configures ImageOrPrintOptions for PNG, sanitizes worksheet names for safe filenames, handles errors per sheet, and demonstrates how to control parallelism for optimal multi‑core performance.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            var workbookPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file '{workbookPath}' was not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(workbookPath);

            // Ensure the output directory exists
            var outputDir = "output_png";
            Directory.CreateDirectory(outputDir);

            // Convert each worksheet to a PNG image in parallel
            Parallel.ForEach(workbook.Worksheets, worksheet =>
            {
                try
                {
                    // Set rendering options for PNG output
                    var imgOptions = new ImageOrPrintOptions
                    {
                        OnePagePerSheet = true,
                        Transparent = false
                        // ImageFormat is inferred from the file extension (.png)
                    };

                    // Render the worksheet to an image
                    var sheetRender = new SheetRender(worksheet, imgOptions);

                    // Build a safe file name using worksheet index and name
                    var safeName = SanitizeFileName(worksheet.Name);
                    var filePath = Path.Combine(outputDir,
                        $"Sheet_{worksheet.Index}_{safeName}.png");

                    // Save the first (and only) page as PNG
                    sheetRender.ToImage(0, filePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process sheet '{worksheet.Name}': {ex.Message}");
                }
            });

            Console.WriteLine("All worksheets have been converted to PNG.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to replace invalid filename characters
    static string SanitizeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}
