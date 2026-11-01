// Title: Convert an Excel workbook loaded from a MemoryStream to a 300 DPI PNG image using Aspose.Cells for .NET
// Common Searches: c# load excel file from byte array and save first sheet as 300 dpi png using Aspose.Cells | how to set horizontal and vertical resolution to 300 DPI when exporting worksheet to PNG with Aspose.Cells | Aspose.Cells render worksheet to high resolution PNG from memory stream in .NET | export Excel worksheet to PNG image with custom DPI in C# Aspose.Cells | convert in‑memory Excel workbook to PNG image using SheetRender and ImageOrPrintOptions
// Tags: memory stream workbook loading Aspose.Cells | ImageOrPrintOptions DPI configuration | SheetRender export worksheet to PNG | high resolution Excel to PNG conversion .NET | one page per sheet image rendering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example reads an Excel file into a byte array, creates a MemoryStream, loads it into an Aspose.Cells Workbook, configures ImageOrPrintOptions for PNG format with 300 DPI horizontal and vertical resolution and one page per sheet, then uses SheetRender to render the first worksheet to a PNG file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.png";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Read the Excel file bytes
            byte[] excelBytes = File.ReadAllBytes(inputPath);

            // Load the workbook from a memory stream
            using (MemoryStream memoryStream = new MemoryStream(excelBytes))
            {
                Workbook workbook = new Workbook(memoryStream);

                // Configure image options: PNG format, 300 DPI resolution
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Default format is PNG; explicit setting omitted to avoid API mismatch
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    OnePagePerSheet = true // Export each sheet as a single image page
                };

                // Render the first worksheet to an image
                SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], imgOptions);
                sheetRender.ToImage(0, outputPath);
                Console.WriteLine($"Worksheet rendered to image: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
