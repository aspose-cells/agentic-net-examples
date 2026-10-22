// Title: Export an Excel worksheet to 1x and 2x PNG files for responsive design using Aspose.Cells for .NET
// AI Prompts: Create C# code that loads a workbook, selects the first worksheet, and saves it as a 96 DPI PNG and a 192 DPI PNG using Aspose.Cells rendering. | Show how to configure ImageOrPrintOptions with different DPI values to produce standard and retina PNG assets from an Excel sheet. | Add robust error handling to a C# Aspose.Cells routine that exports multiple PNG resolutions and logs missing file or rendering exceptions.
// Common Searches: C# Aspose.Cells export worksheet to PNG with custom DPI for retina displays | How to generate low and high resolution PNG images from Excel using Aspose.Cells | Aspose.Cells generate both 96 DPI and 192 DPI PNG files from a worksheet | Responsive web design assets from Excel sheet using Aspose.Cells .NET
// Tags: Aspose.Cells render PNG with custom DPI | C# generate retina PNG from Excel worksheet | custom DPI settings in Aspose.Cells | export Excel to multiple PNG sizes .NET | responsive PNG assets from Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads input.xlsx, accesses the first worksheet, and uses Aspose.Cells SheetRender with ImageOrPrintOptions set to 96 DPI and 192 DPI to create worksheet_1x.png and worksheet_2x.png. It includes file‑existence checks and try‑catch blocks for rendering errors.
class GenerateResponsivePng
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            using (Workbook workbook = new Workbook(inputPath))
            {
                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------- 1x PNG (standard resolution) --------------------
                ImageOrPrintOptions options1x = new ImageOrPrintOptions
                {
                    SaveFormat = SaveFormat.Png,          // Output format
                    HorizontalResolution = 96,           // 1x horizontal DPI
                    VerticalResolution = 96              // 1x vertical DPI
                };

                try
                {
                    // Render and save the 1x PNG
                    SheetRender sr1 = new SheetRender(sheet, options1x);
                    sr1.ToImage(0, "worksheet_1x.png");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error generating 1x PNG: {ex.Message}");
                }

                // -------------------- 2x PNG (high‑resolution) --------------------
                ImageOrPrintOptions options2x = new ImageOrPrintOptions
                {
                    SaveFormat = SaveFormat.Png,          // Output format
                    HorizontalResolution = 192,          // 2x horizontal DPI
                    VerticalResolution = 192             // 2x vertical DPI
                };

                try
                {
                    // Render and save the 2x PNG
                    SheetRender sr2 = new SheetRender(sheet, options2x);
                    sr2.ToImage(0, "worksheet_2x.png");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error generating 2x PNG: {ex.Message}");
                }
            }

            Console.WriteLine("PNG images generated: worksheet_1x.png, worksheet_2x.png");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
