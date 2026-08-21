// Title: Add and rotate a WordArt shape diagonally across an Excel sheet with Aspose.Cells for .NET
// Description: Demonstrates how to insert a WordArt shape using Aspose.Cells, set its dimensions, rotate it 45°, position it from the top‑left to the bottom‑right of a worksheet, and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# | WordArt shape | shape rotation | Excel worksheet | AddWordArt | RotationAngle | diagonal placement | Aspose.Cells example
// Common Searches: Aspose.Cells add WordArt C# | rotate WordArt 45 degrees Aspose.Cells | position WordArt diagonally Excel | set shape rotation angle Aspose.Cells | how to create diagonal header with WordArt in .NET
// Developer Intent: Insert a WordArt object, rotate it 45°, and align it along the worksheet diagonal.
// Use Cases: Create a decorative diagonal banner for automated reports. | Highlight a section of a spreadsheet with rotated branding text. | Generate a watermark‑style WordArt that spans the sheet’s diagonal.
// AI Prompts: Write C# code that uses Aspose.Cells to add a WordArt shape, set its size to 800 × 200 px, rotate it 45°, and anchor it from cell A1 to the opposite corner. | Show how to change the WordArt rotation angle based on a value read from a worksheet cell. | Explain the math for calculating the width and height needed for a WordArt shape to cover the full diagonal of any worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to insert a WordArt shape using Aspose.Cells, set its dimensions, rotate it 45°, position it from the top‑left to the bottom‑right of a worksheet, and save the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset style
        // Positioned at the top‑left corner (row 0, column 0) with a large width to span diagonally
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1,   // preset style
            "Diagonal WordArt",                 // text
            0,   // top row index
            0,   // vertical offset in pixels
            0,   // left column index
            0,   // horizontal offset in pixels
            200, // height in pixels
            800  // width in pixels
        );

        // Rotate the WordArt 45 degrees
        wordArt.RotationAngle = 45;

        // Save the workbook
        workbook.Save("WordArtDiagonal.xlsx");
    }
}
