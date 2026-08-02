// Title: Convert an SVG shape in Aspose.Cells to a PNG image using C#
// Description: Demonstrates how to load an SVG file, add it as a shape to an Aspose.Cells worksheet, and rasterize the sheet to a PNG file with SheetRender and ImageOrPrintOptions. Ideal for exporting Excel content to PNG when SVG support is unavailable.
// Keywords: Aspose.Cells | C# | SVG to PNG conversion | rasterize SVG | SheetRender PNG | ImageOrPrintOptions | Excel PNG export | legacy compatibility | add SVG shape | render worksheet as PNG
// Common Searches: Aspose.Cells convert SVG shape to PNG | C# rasterize SVG in Excel workbook | How to export worksheet with SVG as PNG using Aspose.Cells | SheetRender PNG output for SVG images | Legacy systems SVG to PNG Aspose.Cells .NET
// Developer Intent: The developer needs to transform an SVG shape embedded in an Excel worksheet into a PNG image so that the result can be used by applications that do not support SVG.
// Use Cases: Create PNG reports from Excel files that contain vector logos for older viewers. | Generate thumbnail previews of worksheets with embedded SVG for web portals. | Archive Excel sheets as PNG images when downstream tools cannot render SVG.
// AI Prompts: Provide C# code to rasterize multiple SVG shapes in a worksheet to separate PNG files with Aspose.Cells. | Explain how to control PNG resolution and scaling when rendering SVG shapes using SheetRender. | Show how to replace an SVG shape with its rasterized PNG version inside the same workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgRasterization
{
    // Demonstrates how to load an SVG file, add it as a shape to an Aspose.Cells worksheet, and rasterize the sheet to a PNG file with SheetRender and ImageOrPrintOptions. Ideal for exporting Excel content to PNG when SVG support is unavailable.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Load SVG file bytes (replace with your actual SVG file path)
            byte[] svgBytes;
            using (FileStream fs = new FileStream("sample.svg", FileMode.Open, FileAccess.Read))
            {
                svgBytes = new byte[fs.Length];
                fs.Read(svgBytes, 0, svgBytes.Length);
            }

            // Add the SVG image to the worksheet.
            // Parameters: topRow, top, leftColumn, left, height, width, svgData, compatibleImageData
            // Height and width set to -1 to use original size.
            ShapeCollection shapes = sheet.Shapes;
            shapes.AddSvg(2, 0, 2, 0, -1, -1, svgBytes, null);

            // Configure rendering options to produce a PNG image.
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,   // Output format
                OnePagePerSheet = true       // Render the whole sheet on a single page
            };

            // Create a SheetRender instance with the worksheet and options.
            SheetRender renderer = new SheetRender(sheet, renderOptions);

            // Render the first (and only) page to a PNG file.
            renderer.ToImage(0, "worksheet_rasterized.png");

            // Optional: Save the workbook for reference.
            workbook.Save("worksheet_with_svg.xlsx");

            Console.WriteLine("SVG shape rasterized to PNG successfully.");
        }
    }
}
