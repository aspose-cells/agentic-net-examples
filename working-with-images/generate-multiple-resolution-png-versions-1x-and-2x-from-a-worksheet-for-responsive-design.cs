// Title: Create 1x and 2x PNG files from an Excel worksheet with Aspose.Cells in C# for responsive web images
// AI Prompts: Generate C# code that loads an Excel workbook, picks a worksheet, and renders PNG files at standard (96 dpi) and high‑density (192 dpi) resolutions using Aspose.Cells rendering options. | Write a reusable C# method that takes a Worksheet, an output path, and a DPI value, then saves the sheet as a PNG with a solid white background. | Add robust error handling to a C# Aspose.Cells script that verifies the source XLSX file, creates the destination folder, and logs any failures during multi‑resolution PNG rendering.
// Common Searches: c# aspocells export worksheet to png with 96 dpi and 2x resolution | how to generate retina png from excel sheet using Aspose.Cells | save excel sheet as responsive images for web design c# | Aspose.Cells render first worksheet to png with custom dpi settings | C# example creating multiple resolution pngs from an Excel workbook
// Tags: Aspose.Cells worksheet to PNG with custom DPI | render Excel sheet as high‑density PNG | C# generate responsive PNG images from workbook | image rendering options for Excel to PNG | save worksheet as 1x and 2x PNG files

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an XLSX file, renders the first worksheet to two PNG images—standard 96 dpi (1x) and high‑density 192 dpi (2x)—and saves them in an output folder using Aspose.Cells rendering options.
class GenerateResponsivePng
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string workbookPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Get the first worksheet (or specify by name/index)
            Worksheet sheet = workbook.Worksheets[0];

            // Define and create the output folder
            string outputFolder = "output_images";
            Directory.CreateDirectory(outputFolder);

            // 1x resolution (standard 96 DPI)
            SaveWorksheetAsPng(sheet, Path.Combine(outputFolder, "sheet_1x.png"), 96);

            // 2x resolution (high‑density 192 DPI)
            SaveWorksheetAsPng(sheet, Path.Combine(outputFolder, "sheet_2x.png"), 192);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    /// <param name="sheet">Worksheet to render.</param>
    /// <param name="outputPath">Full path of the PNG file to create.</param>
    /// <param name="dpi">Resolution in dots per inch (e.g., 96 for 1x, 192 for 2x).</param>
    private static void SaveWorksheetAsPng(Worksheet sheet, string outputPath, int dpi)
    {
        try
        {
            // Configure rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                // Set the resolution (DPI) for the output image
                HorizontalResolution = dpi,
                VerticalResolution = dpi,

                // Fit the entire sheet on one page
                OnePagePerSheet = true,

                // Set background to white (transparent by default)
                Transparent = false
            };

            // Render the first page (index 0) to the specified PNG file
            SheetRender sheetRender = new SheetRender(sheet, options);
            sheetRender.ToImage(0, outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to render worksheet to PNG: {ex.Message}");
        }
    }
}
