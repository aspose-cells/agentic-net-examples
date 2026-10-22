// Title: Convert an Excel worksheet with WordArt to a high‑resolution PNG image while preserving gradient fills using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file containing WordArt and exports the first worksheet to a 300 dpi PNG with gradient fill retained using Aspose.Cells. | Show how to configure ImageOrPrintOptions for high‑resolution PNG output and render each sheet to separate image files in C#. | Provide a C# snippet that checks for workbook existence, sets OnePagePerSheet, and saves the worksheet as a transparent‑background PNG with Aspose.Cells.
// Common Searches: how to export Excel sheet with WordArt to PNG using Aspose.Cells C# | preserve gradient fill when converting Excel to high resolution PNG Aspose.Cells | set DPI for PNG export in Aspose.Cells ImageOrPrintOptions C# | render first worksheet as single‑page PNG Aspose.Cells .NET | export WordArt graphics from Excel to PNG without quality loss
// Tags: Aspose.Cells PNG export with gradient fill | ImageOrPrintOptions DPI setting C# | SheetRender ToImage high‑resolution output | WordArt rendering to PNG Aspose.Cells | OnePagePerSheet Excel to image conversion | transparent background PNG export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program verifies the input Excel file, loads it with Aspose.Cells, configures ImageOrPrintOptions for 300 dpi and OnePagePerSheet, then uses SheetRender to export the first worksheet as a high‑resolution PNG that retains WordArt gradient quality.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook that contains WordArt
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet to render
            if (workbook.Worksheets.Count == 0)
                throw new InvalidOperationException("The workbook does not contain any worksheets.");

            // Configure image export options (default format is PNG)
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                // Render each sheet on a single page
                OnePagePerSheet = true,
                // High resolution improves gradient rendering
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Render the first worksheet (index 0) to a PNG image
            SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], options);
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Worksheet rendered successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
