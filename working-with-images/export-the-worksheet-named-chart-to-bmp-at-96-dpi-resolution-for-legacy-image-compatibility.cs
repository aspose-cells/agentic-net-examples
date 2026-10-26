// Title: Export the 'Chart' worksheet to a 96 DPI BMP image using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells export worksheet named Chart to BMP with 96 DPI in C# | C# render Excel sheet to bitmap image at 96 DPI using Aspose.Cells | Save Excel chart worksheet as BMP file using ImageOrPrintOptions 96 DPI | Set horizontal and vertical resolution for BMP export in Aspose.Cells .NET | Export a single page of a worksheet to BMP with SheetRender in C#
// Tags: Aspose.Cells ImageOrPrintOptions BMP resolution | SheetRender export single page bitmap | C# export named worksheet to bitmap | 96 DPI BMP output with Aspose.Cells | Excel chart sheet BMP conversion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing.Imaging;

// The example loads an Excel workbook, selects the worksheet called "Chart", configures ImageOrPrintOptions for BMP format with 96 DPI horizontal and vertical resolution, renders the sheet using SheetRender, and saves the first page as a BMP file while handling missing file and worksheet errors.
class Program
{
    static void Main()
    {
        // Paths for input workbook and output BMP image
        string inputFile = "input.xlsx";
        string outputFile = "Chart.bmp";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
                throw new FileNotFoundException($"Input file not found: {inputFile}");

            // Load the workbook (lifecycle: load)
            Workbook workbook = new Workbook(inputFile);

            // Retrieve the worksheet named "Chart"
            Worksheet chartSheet = workbook.Worksheets["Chart"];
            if (chartSheet == null)
                throw new InvalidOperationException("Worksheet 'Chart' does not exist in the workbook.");

            // Configure image export options (BMP, 96 DPI)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // ImageFormat property is optional; the file extension determines the format.
                HorizontalResolution = 96,   // 96 DPI horizontal
                VerticalResolution = 96,     // 96 DPI vertical
                OnePagePerSheet = true       // Render the whole sheet on one page
            };

            // Render the worksheet to an image (lifecycle: create/render)
            SheetRender sheetRender = new SheetRender(chartSheet, imgOptions);

            // Export the first (and only) page to BMP (lifecycle: save)
            sheetRender.ToImage(0, outputFile);

            Console.WriteLine($"Worksheet exported successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
