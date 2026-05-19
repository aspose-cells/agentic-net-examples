using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddWordArtExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add a WordArt shape using a built‑in preset style
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle5, // Built‑in style: Fill - Gold, Accent 4, Soft Bevel
            "Aspose.Cells WordArt",           // Text to display
            2,    // Upper left row index
            0,    // Vertical offset (pixels) from the top row
            2,    // Upper left column index
            0,    // Horizontal offset (pixels) from the left column
            100,  // Height of the shape (pixels)
            400   // Width of the shape (pixels)
        );

        // Example: set rotation angle (optional)
        wordArt.RotationAngle = 0;

        // Save the workbook to a file
        workbook.Save("WordArtWithPresetStyle.xlsx");
    }
}