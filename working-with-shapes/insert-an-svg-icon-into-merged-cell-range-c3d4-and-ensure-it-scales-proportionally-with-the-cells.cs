// Title: C# – Insert an SVG into merged cells C3:D4 with proportional scaling using Aspose.Cells
// Description: This example creates a workbook, merges the range C3:D4, loads an SVG file (or a fallback SVG string), adds the SVG as a picture anchored to the merged area, locks its aspect ratio, enables the picture to move and resize with the cells, and saves the file as OutputWithSvg.xlsx.
// Keywords: Aspose.Cells SVG insertion | C# add picture to merged cells | lock aspect ratio Aspose.Cells | place SVG in Excel cell | merged range C3:D4 | .NET Excel vector graphic | fallback SVG string | IsPlacedInCell property
// Common Searches: how to add an SVG to a merged cell range with Aspose.Cells .NET | Aspose.Cells keep SVG aspect ratio when resizing cells | place picture inside merged cells C3:D4 Aspose.Cells | load SVG from file with fallback in Aspose.Cells | C# Aspose.Cells example for vector graphics in Excel
// Developer Intent: Add an SVG image to the merged range C3:D4 and have it automatically scale proportionally with the cells.
// Use Cases: Insert a scalable company logo into a merged header for a financial report. | Add responsive icons to a dashboard that adjust when rows or columns are resized. | Create a template that embeds vector graphics in merged cells while preserving quality across PDF and XLSX exports.
// AI Prompts: Write C# code using Aspose.Cells to insert an SVG into a merged cell range and lock its aspect ratio. | Show how to provide a fallback SVG string when the external SVG file is missing while adding the picture to a worksheet. | Explain the effect of the IsPlacedInCell property on picture movement and resizing with merged cells in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, merges the range C3:D4, loads an SVG file (or a fallback SVG string), adds the SVG as a picture anchored to the merged area, locks its aspect ratio, enables the picture to move and resize with the cells, and saves the file as OutputWithSvg.xlsx.
class InsertSvgIntoMergedCells
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge cells C3:D4 (zero‑based indices: row 2, column 2, spanning 2 rows and 2 columns)
            worksheet.Cells.Merge(2, 2, 2, 2);

            // Load SVG file bytes; if the file does not exist, use a simple fallback SVG
            string svgPath = "icon.svg";
            byte[] svgBytes;

            if (File.Exists(svgPath))
            {
                svgBytes = File.ReadAllBytes(svgPath);
            }
            else
            {
                // Minimal SVG content as fallback
                const string fallbackSvg = @"<svg xmlns='http://www.w3.org/2000/svg' width='100' height='100'><rect width='100' height='100' fill='orange'/></svg>";
                svgBytes = System.Text.Encoding.UTF8.GetBytes(fallbackSvg);
                Console.WriteLine($"Warning: '{svgPath}' not found. Using fallback SVG.");
            }

            // Add the SVG as a picture to the merged range.
            // upperLeftRow = 2 (row 3), upperLeftColumn = 2 (column C), offsets = 0.
            ShapeCollection shapes = worksheet.Shapes;
            Picture picture;

            using (MemoryStream ms = new MemoryStream(svgBytes))
            {
                picture = shapes.AddPicture(
                    2, // upperLeftRow
                    2, // upperLeftColumn
                    0, // upperLeftRowOffset
                    0, // upperLeftColumnOffset
                    ms);
            }

            // Lock aspect ratio so the SVG scales proportionally with the cells
            picture.IsAspectRatioLocked = true;

            // Ensure the picture moves/resizes with the merged cells
            picture.IsPlacedInCell = true;

            // Save the workbook
            string outputPath = "OutputWithSvg.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
