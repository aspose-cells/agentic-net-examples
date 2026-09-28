// Title: Render an Excel worksheet to PNG with only cell borders (gridlines hidden) using Aspose.Cells for .NET
// AI Prompts: Generate C# code that turns off worksheet gridlines, adds thin borders to a specific cell range, and saves the sheet as a PNG image using Aspose.Cells. | Show how to set high‑resolution ImageOrPrintOptions and use SheetRender to export a workbook page while preserving only the applied cell borders.
// Common Searches: how to hide gridlines when exporting Excel to PNG with Aspose.Cells C# | Aspose.Cells render worksheet to image with custom borders only | C# code sample that adds thin cell borders and exports the sheet as a PNG image
// Tags: disable worksheet gridlines Aspose.Cells | apply thin borders to range Aspose.Cells | render worksheet to PNG Aspose.Cells | high‑resolution image rendering options | SheetRender export image without gridlines

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

namespace RenderGridlinesExampleApp
{
    // // Demonstrates creating a workbook, applying thin borders to cells A1:B2, disabling gridlines, configuring image options, and rendering the worksheet to a PNG image using Aspose.Cells.
    class RenderGridlinesExample
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Header 1");
                sheet.Cells["B1"].PutValue("Header 2");
                sheet.Cells["A2"].PutValue("Data 1");
                sheet.Cells["B2"].PutValue("Data 2");

                // Apply borders so cells are visible when gridlines are hidden
                Style borderStyle = workbook.CreateStyle();
                borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

                // Apply the style to the range A1:B2
                AsposeRange range = sheet.Cells.CreateRange("A1:B2");
                range.ApplyStyle(borderStyle, new StyleFlag { Borders = true });

                // Hide worksheet gridlines
                sheet.IsGridlinesVisible = false;

                // Configure rendering options (gridlines already hidden)
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    HorizontalResolution = 150,
                    VerticalResolution = 150
                };

                // Render the worksheet to an image
                SheetRender sr = new SheetRender(sheet, renderOptions);
                string outputPath = "RenderedSheet.png";

                // Ensure any existing file is overwritten safely
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                sr.ToImage(0, outputPath);

                // Verify that the file was created
                if (File.Exists(outputPath))
                {
                    Console.WriteLine($"Rendering completed. Check '{outputPath}' for output without solid gridlines.");
                }
                else
                {
                    Console.WriteLine("Rendering completed, but the output file was not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
