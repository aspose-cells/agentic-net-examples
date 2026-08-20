// Title: Export Excel Worksheet to 1x & 2x PNG with Aspose.Cells for .NET – Retina‑Ready Images
// Description: Loads an Excel workbook, validates the source file, and uses Aspose.Cells SheetRender with ImageOrPrintOptions to create two PNG files from the first worksheet page: a standard‑resolution (96 dpi) image for 1x displays and a high‑resolution (192 dpi) image for 2x/retina screens, ideal for responsive web design.
// Keywords: Aspose.Cells PNG export | C# Excel to PNG | multiple DPI image export | retina ready Excel image | responsive web images from Excel | SheetRender high resolution | ImageOrPrintOptions DPI | 1x 2x PNG Aspose.Cells | .NET Excel image rendering
// Common Searches: export Excel worksheet as 2x PNG Aspose.Cells | Aspose.Cells generate retina PNG from Excel | set DPI for PNG export using Aspose.Cells .NET | create 1x and 2x images from Excel sheet | responsive image export Aspose.Cells C#
// Developer Intent: Create both standard‑resolution (1x) and high‑resolution (2x) PNG files from a worksheet for use in responsive layouts.
// Use Cases: Produce a low‑resolution thumbnail and a retina‑ready PNG for HTML srcset. | Automate web‑ready image assets from Excel reports, delivering 1x and 2x files in one run. | Embed worksheet visuals in mobile and desktop pages with appropriate DPI handling.
// AI Prompts: Generate C# code with Aspose.Cells to export a worksheet to PNG at 72 dpi and 144 dpi. | Show how to modify the sample to export every worksheet in a workbook to both 1x and 2x PNG files. | Give best‑practice error handling for rendering large worksheets to high‑resolution PNG using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace ResponsiveImageExport
{
    // Loads an Excel workbook, validates the source file, and uses Aspose.Cells SheetRender with ImageOrPrintOptions to create two PNG files from the first worksheet page: a standard‑resolution (96 dpi) image for 1x displays and a high‑resolution (192 dpi) image for 2x/retina screens, ideal for responsive web design.
    public class ExportMultipleResolutions
    {
        public static void Run()
        {
            const string sourcePath = "SourceWorkbook.xlsx";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
                return;
            }

            try
            {
                // Load the source workbook
                Workbook workbook = new Workbook(sourcePath);
                Worksheet worksheet = workbook.Worksheets[0];

                // ---------- 1x resolution (default 96 DPI) ----------
                ImageOrPrintOptions options1x = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true,
                    HorizontalResolution = 96,
                    VerticalResolution = 96
                    // Default image format is PNG, no need to set explicitly
                };

                // Render the first page to a PNG file at 1x resolution
                SheetRender render1x = new SheetRender(worksheet, options1x);
                render1x.ToImage(0, "Worksheet_1x.png");

                // ---------- 2x resolution (192 DPI) ----------
                ImageOrPrintOptions options2x = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true,
                    HorizontalResolution = 192,
                    VerticalResolution = 192
                    // Default image format is PNG
                };

                // Render the same page to a higher‑resolution PNG
                SheetRender render2x = new SheetRender(worksheet, options2x);
                render2x.ToImage(0, "Worksheet_2x.png");

                Console.WriteLine("Export completed: 1x and 2x PNG images generated.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred during export: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportMultipleResolutions.Run();
        }
    }
}
