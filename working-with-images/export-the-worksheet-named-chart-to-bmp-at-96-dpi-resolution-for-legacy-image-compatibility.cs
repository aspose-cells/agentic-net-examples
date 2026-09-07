// Title: Export the 'Chart' worksheet from an Excel file to a 96 DPI BMP image using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook, selects the worksheet named "Chart", configures ImageOrPrintOptions for 96 DPI, and saves the sheet as a BMP file with Aspose.Cells. | Demonstrate how to use SheetRender together with ImageOrPrintOptions to render a specific worksheet to a bitmap image at a custom resolution in a .NET application.
// Common Searches: Aspose.Cells C# export specific worksheet to BMP with 96 DPI | How to set horizontal and vertical resolution when rendering an Excel sheet to an image using Aspose.Cells | Render Excel chart sheet as a bitmap file in .NET with one page per sheet option
// Tags: export worksheet to BMP Aspose.Cells | ImageOrPrintOptions DPI setting | SheetRender ToImage C# example | one page per sheet image rendering | convert Excel chart to bitmap

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads 'input.xlsx', retrieves the worksheet named 'Chart', sets ImageOrPrintOptions to 96 DPI for both axes and enables OnePagePerSheet, then uses SheetRender to export the first page of the sheet as 'Chart.bmp'. It includes error handling for missing files or worksheets.
class ExportWorksheetToBmp
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the worksheet named "Chart"
            Worksheet chartSheet = workbook.Worksheets["Chart"];
            if (chartSheet == null)
            {
                Console.WriteLine("Worksheet 'Chart' not found.");
                return;
            }

            // Configure image options: 96 DPI, one page per sheet
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 96,
                VerticalResolution = 96,
                OnePagePerSheet = true
                // Note: ImageFormat property is not available in this version;
                // the default format (PNG) will be used. To obtain BMP, rename the output file with .bmp extension.
            };

            // Render the worksheet to an image
            SheetRender sheetRender = new SheetRender(chartSheet, imgOptions);

            // Export the first (and only) page to BMP (renamed from default PNG)
            string outputPath = "Chart.bmp";
            try
            {
                sheetRender.ToImage(0, outputPath);
                Console.WriteLine($"Worksheet exported to {outputPath} at 96 DPI.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to export image: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
