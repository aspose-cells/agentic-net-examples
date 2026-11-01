// Title: Export a specific worksheet to a PNG image at 300 DPI with a transparent background using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export Sheet1 to PNG with 300 DPI and transparent background | How to render an Excel worksheet as a high‑resolution PNG using Aspose.Cells .NET | Set transparent background when saving an Excel sheet as an image with Aspose.Cells | ImageOrPrintOptions 300 DPI transparent PNG example in C# | Export a single Excel worksheet to PNG file using SheetRender in C#
// Tags: export worksheet to PNG with 300 DPI Aspose.Cells | transparent background image export ImageOrPrintOptions | SheetRender render Excel sheet as high‑resolution PNG | Aspose.Cells C# save specific sheet as PNG image | set horizontal and vertical resolution 300 DPI Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies the input Excel file, loads it into a Workbook, selects a specific worksheet, configures ImageOrPrintOptions to enable a transparent background and set both horizontal and vertical resolutions to 300 DPI, then uses SheetRender to export the worksheet as a single PNG image to the designated output path, with error handling for missing files or runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.png";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Select the worksheet to export (by name or index)
            Worksheet sheet = workbook.Worksheets["Sheet1"]; // change name or use workbook.Worksheets[0] as needed

            // Configure image export options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // ImageFormat property removed because it may not be available in some versions;
                // the default format is PNG, which supports transparency.
                OnePagePerSheet = true,   // Export the whole sheet as one image
                Transparent = true,       // Enable transparent background
                HorizontalResolution = 300, // 300 DPI horizontal
                VerticalResolution = 300    // 300 DPI vertical
            };

            // Render the worksheet to an image using the specified options
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);

            // Export the first (and only) page to a PNG file
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Worksheet exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
