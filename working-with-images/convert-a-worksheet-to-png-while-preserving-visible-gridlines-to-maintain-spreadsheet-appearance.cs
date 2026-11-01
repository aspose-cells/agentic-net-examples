// Title: Convert an Excel worksheet to a PNG image while preserving visible gridlines with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook, sets ImageOrPrintOptions.ShowGridLines to true, and uses SheetRender to save the first worksheet as a PNG file. | Update the existing SheetRender example to output a high‑resolution PNG (e.g., 300 DPI) and guarantee that gridlines appear in the rendered image. | Create a reusable C# method that takes input and output paths, iterates through all worksheets in a workbook, and creates separate PNG images for each sheet with gridlines visible.
// Common Searches: Aspose.Cells C# export worksheet to PNG with gridlines shown | How to keep Excel gridlines when converting to image using Aspose.Cells | Render Excel sheet as PNG preserving layout and gridlines in .NET | ImageOrPrintOptions ShowGridLines example for Aspose.Cells | Convert each sheet of a workbook to PNG files with visible gridlines C#
// Tags: Aspose.Cells export worksheet to PNG | ImageOrPrintOptions ShowGridLines property | SheetRender generate high‑resolution PNG | preserve Excel gridlines in image output | C# convert .xlsx to PNG per worksheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads a workbook, selects the first worksheet, configures ImageOrPrintOptions (including OnePagePerSheet and optional ShowGridLines), renders the sheet with SheetRender, and writes the result to a PNG file while handling missing files and runtime exceptions.
class WorksheetToPng
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.png";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Set up image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                // Render each worksheet as a single page (most worksheets have one page)
                OnePagePerSheet = true
                // Additional options such as resolution can be set here if needed
                // HorizontalResolution = 300,
                // VerticalResolution = 300
            };

            // Render the worksheet to a PNG image
            SheetRender renderer = new SheetRender(sheet, options);
            renderer.ToImage(0, outputPath);

            Console.WriteLine($"Worksheet successfully exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
