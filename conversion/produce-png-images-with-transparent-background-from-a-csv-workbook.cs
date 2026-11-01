// Title: Generate PNG images with transparent background from a CSV workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export CSV worksheet to transparent PNG file | How to render CSV data as PNG with alpha channel using Aspose.Cells | Aspose.Cells ImageOrPrintOptions Transparent true for CSV to PNG conversion | Save each sheet of a CSV workbook as separate transparent PNG in .NET | Load CSV with LoadOptions and create PNG images with transparency in C#
// Tags: Aspose.Cells CSV to PNG conversion | ImageOrPrintOptions transparent background | SheetRender PNG export per worksheet | OnePagePerSheet image rendering | LoadOptions CSV format loading

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads a CSV file as a Workbook using LoadOptions, sets ImageOrPrintOptions to enable transparent backgrounds and one-page-per-sheet rendering, then iterates through each worksheet and uses SheetRender to save each sheet as an individual PNG file with an alpha channel.
class CsvToPngTransparent
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Verify that the CSV file exists to avoid FileNotFoundException
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"Error: The file \"{csvPath}\" was not found.");
            return;
        }

        try
        {
            // Load the CSV workbook using LoadOptions to specify CSV format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Configure image rendering options: PNG format with transparent background
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // PNG is the default format; no need to set ImageFormat explicitly
                Transparent = true,          // Enable transparent background
                OnePagePerSheet = true       // Render each sheet on a single page
            };

            // Render each worksheet to a separate PNG file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                SheetRender renderer = new SheetRender(sheet, imgOptions);

                // Build output file name (e.g., input_Sheet1.png, input_Sheet2.png, ...)
                string outputFileName = $"{Path.GetFileNameWithoutExtension(csvPath)}_Sheet{i + 1}.png";

                try
                {
                    // Render the first page of the worksheet to an image and save it
                    renderer.ToImage(0, outputFileName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to render sheet {i + 1}: {ex.Message}");
                }
            }

            Console.WriteLine("PNG images with transparent background have been generated.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
