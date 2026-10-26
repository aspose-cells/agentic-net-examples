// Title: Render a worksheet by zero‑based index to a high‑resolution PNG image at 200 DPI using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# render specific sheet index to PNG with custom DPI | how to set 200 DPI when converting Excel worksheet to image using Aspose.Cells | C# SheetRender export first worksheet as high‑resolution PNG | ImageOrPrintOptions horizontal and vertical resolution settings for Excel to image conversion | save Excel worksheet as detailed PNG file using Aspose.Cells .NET API
// Tags: worksheet index to PNG Aspose.Cells | 200 DPI image export Excel | SheetRender high resolution PNG | ImageOrPrintOptions resolution configuration | C# export specific sheet as image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example checks for the input Excel file, loads it into a Workbook, selects a worksheet by its zero‑based index, configures ImageOrPrintOptions with 200 dpi horizontal and vertical resolution, and uses SheetRender to render the first page of that worksheet to a PNG file, handling any loading or rendering errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.png";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Index of the worksheet to render (0‑based)
        int sheetIndex = 0;
        Worksheet worksheet = workbook.Worksheets[sheetIndex];

        // Configure image options: 200 DPI resolution (default format is PNG)
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            HorizontalResolution = 200,
            VerticalResolution = 200
        };

        try
        {
            // Render the worksheet to an image (page 0)
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
            sheetRender.ToImage(0, outputPath);
            Console.WriteLine($"Worksheet rendered successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during rendering: {ex.Message}");
        }
    }
}
