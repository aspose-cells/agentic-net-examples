// Title: Insert SVG into merged cells C3:D4 with proportional scaling using Aspose.Cells for .NET
// Description: Demonstrates how to merge cells C3:D4, load an SVG file, add it as a picture that fills the merged range, lock its aspect ratio, set the picture to move and size with the cells, and save the workbook as an XLSX file with Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | SVG insertion | merged cells | picture placement | aspect ratio lock | MoveAndSize | AddPicture | Excel workbook | vector graphics
// Common Searches: Aspose.Cells add SVG to merged cells | C# lock picture aspect ratio in Excel | Place vector icon in merged range Aspose.Cells | MoveAndSize placement for shapes .NET | How to scale SVG with merged cells in Excel
// Developer Intent: Embed an SVG image into a merged cell range and have it resize proportionally with the cells.
// Use Cases: Add a scalable company logo to a header that spans merged cells. | Insert vector icons into dashboard worksheets where column widths may change. | Create printable reports with SVG graphics that maintain quality when cells are resized.
// AI Prompts: Write C# code with Aspose.Cells to insert an SVG into a merged cell range and lock its aspect ratio. | Show how to configure the Placement property so a picture moves and sizes with merged cells in Aspose.Cells. | Provide error handling for missing SVG files when adding them as pictures to a worksheet.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to merge cells C3:D4, load an SVG file, add it as a picture that fills the merged range, lock its aspect ratio, set the picture to move and size with the cells, and save the workbook as an XLSX file with Aspose.Cells for C#.
class InsertSvgIntoMergedCells
{
    static void Main()
    {
        try
        {
            // Path to the SVG file
            string svgPath = "icon.svg";

            // Verify that the SVG file exists before attempting to read it
            if (!File.Exists(svgPath))
            {
                Console.WriteLine($"SVG file not found: {svgPath}");
                return;
            }

            // Load SVG file bytes
            byte[] svgBytes = File.ReadAllBytes(svgPath);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge cells C3:D4 (zero‑based indices: row 2, column 2, spanning 2 rows and 2 columns)
            worksheet.Cells.Merge(2, 2, 2, 2);

            // Add the SVG as a picture to the merged range.
            // The picture will occupy the merged cells (rows 2‑3, columns 2‑3).
            ShapeCollection shapes = worksheet.Shapes;
            using (MemoryStream ms = new MemoryStream(svgBytes))
            {
                Picture picture = shapes.AddPicture(2, 2, 3, 3, ms);
                // Lock aspect ratio so the SVG scales proportionally with the merged cells
                picture.IsAspectRatioLocked = true;
                // Set the picture to move and size with cells
                picture.Placement = PlacementType.MoveAndSize;
            }

            // Save the workbook
            string outputPath = "OutputWithSvg.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
