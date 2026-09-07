// Title: Convert a specific Excel worksheet to a PNG image while keeping Excel gridlines visible using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, selects a worksheet, and saves it as a PNG image with gridlines displayed using Aspose.Cells. | Demonstrate how to set ImageOrPrintOptions (e.g., GridlineType, OnePagePerSheet) to render an Excel sheet to PNG while preserving the original gridline appearance in .NET.
// Common Searches: C# Aspose.Cells export worksheet to PNG with gridlines | How to keep Excel gridlines when rendering to image using Aspose.Cells | ImageOrPrintOptions GridlineType example for PNG output | Render Excel sheet as PNG on a single page using Aspose.Cells .NET | Save first worksheet of workbook as PNG preserving cell borders in C#
// Tags: Aspose.Cells worksheet PNG export preserving gridlines | ImageOrPrintOptions GridlineType configuration | OnePagePerSheet rendering option for PNG output | C# convert Excel sheet to image with visible gridlines | Aspose.Cells image rendering settings for Excel appearance

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample checks for the input file, loads it into an Aspose.Cells Workbook, selects the first worksheet, configures ImageOrPrintOptions to render the sheet on a single page with gridlines visible, creates a SheetRender, and saves the result as a PNG image while handling any exceptions.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify by name/index)
            Worksheet sheet = workbook.Worksheets[0];

            // Configure image export options
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                // Render the whole sheet on one page
                OnePagePerSheet = true
                // Default image format is PNG; no need to set explicitly for compatibility
            };

            // Render the worksheet to a PNG image
            SheetRender renderer = new SheetRender(sheet, options);
            renderer.ToImage(0, outputPath);

            Console.WriteLine($"Worksheet rendered successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
