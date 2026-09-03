// Title: Export the first worksheet page to a 150 DPI PNG image using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells and saves the first worksheet as a PNG image at 150 DPI. | Show how to configure ImageOrPrintOptions for horizontal and vertical resolution and use SheetRender to generate a PNG file. | Add checks for the source file’s existence and create the destination directory before exporting the worksheet to PNG.
// Common Searches: Aspose.Cells C# export worksheet page to PNG with specific DPI | how to set 150 dpi when converting Excel sheet to image using Aspose.Cells | C# example for saving Excel worksheet as high‑resolution PNG | render first sheet of workbook to PNG file with custom resolution Aspose.Cells
// Tags: worksheet to PNG conversion with custom DPI Aspose.Cells | ImageOrPrintOptions DPI setting C# | SheetRender export Excel sheet as PNG | C# file existence check before Aspose.Cells export | create output directory for Aspose.Cells image rendering

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.IO;

// The example loads an Excel workbook, verifies the input file, configures ImageOrPrintOptions to 150 DPI, uses SheetRender to render the first worksheet page, creates the output folder if missing, and saves the result as a PNG image.
class Program
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
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or any specific worksheet you need)
            Worksheet sheet = workbook.Worksheets[0];

            // Configure image rendering options: 150 DPI resolution
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Image format defaults to PNG when the file extension is .png,
                // so we omit setting ImageFormat to avoid API version issues.
                HorizontalResolution = 150,
                VerticalResolution = 150
            };

            // Create a SheetRender object using the worksheet and the image options
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the first page of the worksheet to a PNG file
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Worksheet exported successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
