// Title: Add WordArt text at specific point coordinates and export to PDF with Aspose.Cells for .NET
// AI Prompts: Create a free‑floating TextEffect shape, assign its Left, Top, Width, and Height in points, and save the workbook as a PDF. | Position a WordArt shape at exact page coordinates (e.g., 150 pt left, 200 pt top) before converting the worksheet to PDF using Aspose.Cells. | Generate a PDF where custom text appears at precise locations by setting absolute coordinates on a shape in C#.
// Common Searches: Aspose.Cells C# set WordArt shape coordinates before PDF export | how to use free floating placement for text effect in Aspose.Cells PDF | position custom text at exact points in generated PDF with Aspose.Cells | C# Aspose.Cells place shape at (150,200) points and save as PDF | absolute positioning of shapes in Aspose.Cells worksheet for PDF output
// Tags: Aspose.Cells free‑floating shape positioning | C# set shape left top points | WordArt absolute coordinates PDF | text effect shape PDF export Aspose.Cells | custom text placement PDF Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds a free‑floating WordArt (TextEffect) shape, sets its Left, Top, Width, and Height properties in points to position it precisely on the page, and saves the workbook as a PDF, resulting in custom text placed at exact coordinates.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a text effect shape (WordArt) to the worksheet
                // Parameters: preset effect, text, font name, font size, bold, italic,
                // upper left row, upper left column, lower right row, lower right column,
                // shape width, shape height
                Shape textShape = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    "Custom Text placed precisely",
                    "Arial",
                    14,
                    false,
                    false,
                    0,   // upper left row (cell based, will be overridden by absolute positioning)
                    0,   // upper left column
                    0,   // lower right row
                    0,   // lower right column
                    0,   // shape width (will be set later)
                    0    // shape height (will be set later)
                );

                // Set the shape to free‑floating so we can position it with absolute coordinates (points)
                textShape.Placement = PlacementType.FreeFloating;

                // Define the exact position on the page (in points)
                // 1 point = 1/72 inch.
                textShape.Left = 150;   // distance from the left edge of the page
                textShape.Top = 200;    // distance from the top edge of the page

                // Optionally set the size of the shape (also in points)
                textShape.Width = 300;
                textShape.Height = 50;

                // Save the workbook as a PDF; the text will appear at the specified coordinates
                workbook.Save("CustomText.pdf", SaveFormat.Pdf);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
