// Title: Convert CSV Workbook to Transparent PNG Images with Aspose.Cells for .NET (C#)
// Description: Loads a CSV file into an Aspose.Cells Workbook, configures ImageOrPrintOptions with Transparent=true, and uses SheetRender to export each worksheet (or page) as a PNG file that has a transparent background. Ideal for web overlays, dashboard graphics, and batch image generation.
// Keywords: Aspose.Cells | C# | .NET | CSV to PNG | transparent PNG | ImageOrPrintOptions | SheetRender | export worksheet as image | batch conversion | image rendering
// Common Searches: Aspose.Cells transparent PNG from CSV | C# export CSV to PNG with transparency | how to render worksheet as PNG with transparent background | save Aspose.Cells workbook as PNG image transparent | batch convert CSV files to PNG using Aspose.Cells
// Developer Intent: Generate PNG images with a transparent background from a CSV workbook using Aspose.Cells in C#.
// Use Cases: Overlay CSV data on web pages where the page background must remain visible. | Create theme‑aware PNG assets for reporting dashboards or mobile apps. | Automate batch conversion of multiple CSV files into transparent PNG files for UI integration. | Produce image tiles without background for printable or PDF reports.
// AI Prompts: Write C# code that reads a CSV file into an Aspose.Cells Workbook and saves each worksheet as a transparent PNG image. | Show how to set ImageOrPrintOptions.Transparent to true and render a worksheet to PNG with Aspose.Cells. | Explain how to modify the sample to add a custom background color instead of transparency. | Provide a script that processes all CSV files in a folder and outputs transparent PNGs for each workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Loads a CSV file into an Aspose.Cells Workbook, configures ImageOrPrintOptions with Transparent=true, and uses SheetRender to export each worksheet (or page) as a PNG file that has a transparent background. Ideal for web overlays, dashboard graphics, and batch image generation.
class CsvToTransparentPng
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "data.csv";

        // Directory where PNG images will be saved
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Import CSV data into the first worksheet starting at cell A1
        // Using comma as delimiter and converting numeric strings to numbers
        workbook.Worksheets[0].Cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Configure image rendering options for transparent PNG output
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,   // PNG format
            Transparent = true,          // Enable transparent background
            OnePagePerSheet = true       // Render each sheet as a single page
        };

        // Iterate through all worksheets and render each to a PNG file
        for (int sheetIdx = 0; sheetIdx < workbook.Worksheets.Count; sheetIdx++)
        {
            Worksheet sheet = workbook.Worksheets[sheetIdx];

            // Create a SheetRender instance with the worksheet and image options
            SheetRender renderer = new SheetRender(sheet, imgOptions);

            // Render each page of the sheet (usually one page because of OnePagePerSheet)
            for (int page = 0; page < renderer.PageCount; page++)
            {
                string fileName = Path.Combine(outputDir,
                    $"Sheet{sheetIdx + 1}_Page{page + 1}.png");

                // Save the rendered page as a PNG file with transparent background
                renderer.ToImage(page, fileName);
            }
        }

        Console.WriteLine("Transparent PNG images have been generated successfully.");
    }
}
