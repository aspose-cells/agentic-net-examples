// Title: C# – Insert an SVG (logo.svg) into cells E5:F6 while keeping its original size with Aspose.Cells
// Description: Creates a new workbook, loads logo.svg as a byte array, adds it with AddSvg (height = -1, width = -1) to preserve native dimensions, then sets UpperLeftRow/Column and LowerRightRow/Column so the shape occupies the E5:F6 range, and saves the file as output.xlsx.
// Keywords: Aspose.Cells C# | AddSvg | insert SVG into Excel | preserve SVG dimensions | shape placement E5 F6 | logo.svg worksheet | Aspose.Cells Drawing | Excel SVG example | C# Excel automation
// Common Searches: how to add an SVG to a specific cell range using Aspose.Cells | preserve original size of SVG when inserting into Excel C# | Aspose.Cells place image in cells E5 to F6 | C# code to embed logo.svg in a workbook with Aspose | Aspose.Cells shape positioning by row and column
// Developer Intent: Place logo.svg into the worksheet at E5:F6 without scaling the graphic.
// Use Cases: Add a scalable company logo to a report header that spans two cells. | Embed vector graphics in generated invoices while retaining crisp quality. | Create a template that positions an SVG watermark inside a defined cell block.
// AI Prompts: Show how to center the SVG inside the E5:F6 range while preserving its native size with Aspose.Cells. | Provide a loop example that inserts multiple SVG files into different cell ranges in a .NET workbook. | Explain how to detect an oversized SVG and automatically scale it to fit a target range using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace InsertSvgExampleApp
{
    // Creates a new workbook, loads logo.svg as a byte array, adds it with AddSvg (height = -1, width = -1) to preserve native dimensions, then sets UpperLeftRow/Column and LowerRightRow/Column so the shape occupies the E5:F6 range, and saves the file as output.xlsx.
    class InsertSvgExample
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the SVG file
                string svgPath = "logo.svg";

                // Ensure the SVG file exists before loading
                if (!File.Exists(svgPath))
                {
                    Console.WriteLine($"SVG file not found: {svgPath}");
                    return;
                }

                // Load the SVG file into a byte array
                byte[] svgBytes = File.ReadAllBytes(svgPath);

                // Add the SVG to the worksheet at cell E5 (row 4, column 4)
                ShapeCollection shapes = worksheet.Shapes;
                Picture picture = shapes.AddSvg(
                    topRow: 4,
                    top: 0,
                    leftColumn: 4,
                    left: 0,
                    height: -1,
                    width: -1,
                    svgData: svgBytes,
                    compatibleImageData: null);

                // Adjust placement to fit within the target range (E5:F6)
                picture.UpperLeftRow = 4;   // E5
                picture.UpperLeftColumn = 4; // E
                picture.LowerRightRow = 5;   // F6
                picture.LowerRightColumn = 5; // F

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
