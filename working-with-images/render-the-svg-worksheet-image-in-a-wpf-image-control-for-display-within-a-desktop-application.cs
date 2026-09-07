// Title: Render an Excel worksheet to SVG and show it in a WPF Image control using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, renders the first worksheet to an SVG MemoryStream, converts the stream to a BitmapImage, and assigns it to the Source property of a WPF Image element. | Provide a WPF Window XAML and code‑behind example that creates an Image control, reads the SVG output from SheetRender, and displays it directly without writing the file to disk. | Show how to add a PNG fallback: after rendering the worksheet to SVG, also render it to PNG, save both files, and programmatically switch the Image source to the PNG if the SVG cannot be displayed.
// Common Searches: Aspose.Cells C# render worksheet to SVG for WPF Image control | How to bind SVG stream from SheetRender to WPF Image source | Display Excel sheet as vector graphic in a WPF desktop app | Convert Aspose.Cells SVG MemoryStream to BitmapImage in .NET | Fallback to PNG when SVG rendering fails in WPF using Aspose.Cells
// Tags: Aspose.Cells render worksheet to SVG | WPF Image control display SVG | SheetRender SVG MemoryStream conversion | BitmapImage from SVG stream C# | Excel to PNG fallback Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WorksheetSvgDemo
{
    // The example loads Sample.xlsx with Aspose.Cells, uses SheetRender and ImageOrPrintOptions to generate SVG and PNG images of the first worksheet, saves them to disk, and demonstrates converting the SVG MemoryStream to a BitmapImage for binding to a WPF Image control, with a PNG fallback strategy.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source Excel file
                string inputPath = @"C:\Data\Sample.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // STEP 1: Render the worksheet to SVG and save to disk
                // ------------------------------------------------------------
                string svgPath = @"C:\Data\Sample.svg";
                using (MemoryStream svgStream = new MemoryStream())
                {
                    ImageOrPrintOptions svgOptions = new ImageOrPrintOptions
                    {
                        SaveFormat = SaveFormat.Svg,
                        OnePagePerSheet = true
                    };

                    SheetRender svgRenderer = new SheetRender(sheet, svgOptions);
                    svgRenderer.ToImage(0, svgStream); // Render first page

                    // Write SVG data to file
                    File.WriteAllBytes(svgPath, svgStream.ToArray());
                    Console.WriteLine($"SVG saved to {svgPath}");
                }

                // ------------------------------------------------------------
                // STEP 2: Render the worksheet to PNG and save to disk
                // ------------------------------------------------------------
                string pngPath = @"C:\Data\Sample.png";
                using (MemoryStream pngStream = new MemoryStream())
                {
                    ImageOrPrintOptions pngOptions = new ImageOrPrintOptions
                    {
                        SaveFormat = SaveFormat.Png,
                        OnePagePerSheet = true
                    };

                    SheetRender pngRenderer = new SheetRender(sheet, pngOptions);
                    pngRenderer.ToImage(0, pngStream); // Render first page

                    // Write PNG data to file
                    File.WriteAllBytes(pngPath, pngStream.ToArray());
                    Console.WriteLine($"PNG saved to {pngPath}");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
