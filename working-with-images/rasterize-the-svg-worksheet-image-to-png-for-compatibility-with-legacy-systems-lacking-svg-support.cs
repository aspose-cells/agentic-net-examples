// Title: Rasterize Worksheet SVG to PNG with Aspose.Cells for .NET
// Description: Demonstrates how to load an external SVG file, embed it in a worksheet using AddSvg, configure SvgImageOptions to output PNG, rasterize the SVG with the ToImage method into a MemoryStream, and save the PNG file while optionally preserving the original SVG in the workbook—ideal for legacy systems that do not support SVG.
// Keywords: Aspose.Cells SVG to PNG | rasterize SVG worksheet | SvgImageOptions ImageType Png | AddSvg C# example | export SVG as PNG Aspose.Cells | convert SVG shape to raster image | legacy system SVG compatibility
// Common Searches: Aspose.Cells convert SVG to PNG | How to rasterize SVG in a worksheet .NET | SvgImageOptions ImageType.Png example | AddSvg and ToImage usage | C# export worksheet SVG as PNG
// Developer Intent: Generate a PNG version of an SVG image placed in an Excel worksheet to support environments that require raster formats.
// Use Cases: Create PNG thumbnails of SVG charts embedded in workbooks for reporting dashboards. | Batch‑process multiple SVG shapes across worksheets and deliver PNG streams to web APIs. | Save a PNG copy of an SVG logo while keeping the original SVG in the Excel file for future editing.
// AI Prompts: Write C# code that loads an SVG file, adds it to an Aspose.Cells worksheet with AddSvg, and saves it as a PNG using SvgImageOptions. | Explain step‑by‑step how to rasterize an SVG shape in a workbook to a MemoryStream and write the PNG to disk, including error handling. | Show how to loop through all worksheets in a workbook, find SVG pictures, and export each to a separate PNG file with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace SvgToPngExample
{
    // Demonstrates how to load an external SVG file, embed it in a worksheet using AddSvg, configure SvgImageOptions to output PNG, rasterize the SVG with the ToImage method into a MemoryStream, and save the PNG file while optionally preserving the original SVG in the workbook—ideal for legacy systems that do not support SVG.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                string svgPath = "input.svg";
                if (!File.Exists(svgPath))
                {
                    Console.WriteLine($"SVG file not found: {svgPath}");
                    return;
                }

                // Load SVG file bytes
                byte[] svgData = File.ReadAllBytes(svgPath);

                // Add the SVG image to the worksheet
                // Parameters: upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, height, width, svgData, compatibleImageData
                Picture svgPicture = worksheet.Shapes.AddSvg(0, 0, 0, 0, -1, -1, svgData, null);

                // Prepare SVG rendering options to rasterize to PNG
                SvgImageOptions pngOptions = new SvgImageOptions
                {
                    ImageType = ImageType.Png // Rasterize SVG to PNG
                };

                // Render the SVG picture to a PNG image using a memory stream
                using (MemoryStream pngStream = new MemoryStream())
                {
                    svgPicture.ToImage(pngStream, pngOptions);
                    File.WriteAllBytes("output.png", pngStream.ToArray());
                }

                // Save the workbook (optional, shows the SVG remains in the sheet)
                workbook.Save("workbook_with_svg.xlsx");
                Console.WriteLine("Processing completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
