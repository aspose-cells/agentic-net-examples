// Title: Add and rotate a WordArt shape diagonally in an Excel sheet using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a WordArt shape with a preset style, sizes it to 500 × 500 px, rotates it 45°, positions it from the top‑left to the bottom‑right corner, and saves the file as WordArtDiagonal.xlsx.
// Keywords: Aspose.Cells C# | WordArt shape | add WordArt Excel | rotate shape Aspose.Cells | diagonal placement WordArt | PresetWordArtStyle | RotationAngle property | Excel worksheet shapes | .NET Excel graphics
// Common Searches: how to add WordArt with Aspose.Cells C# | rotate WordArt shape 45 degrees in Excel using Aspose | position WordArt diagonally across a worksheet | Aspose.Cells preset WordArt style example | C# code to create tilted WordArt in Excel
// Developer Intent: Insert a WordArt object, tilt it 45°, and span it across the worksheet diagonal.
// Use Cases: Create a slanted banner for report titles. | Add a watermark‑style WordArt that runs corner‑to‑corner. | Design decorative headings for presentation‑style workbooks.
// AI Prompts: Generate C# Aspose.Cells code that adds a preset WordArt shape, sets its size, rotates it 45°, and saves the workbook. | Show how to place a WordArt object diagonally across an Excel sheet and adjust its rotation with Aspose.Cells for .NET. | Explain how to compute shape dimensions so a WordArt covers the full diagonal of any worksheet using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a WordArt shape with a preset style, sizes it to 500 × 500 px, rotates it 45°, positions it from the top‑left to the bottom‑right corner, and saves the file as WordArtDiagonal.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the shapes collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add a WordArt shape with a preset style
        // Positioned at the top‑left corner and sized to span diagonally across the sheet
        Shape wordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset style
            "Diagonal WordArt",               // text
            0,    // topRow index
            0,    // top offset (pixels)
            0,    // leftColumn index
            0,    // left offset (pixels)
            500,  // height (pixels)
            500   // width (pixels)
        );

        // Rotate the WordArt shape by 45 degrees
        wordArt.RotationAngle = 45;

        // Save the workbook to a file
        workbook.Save("WordArtDiagonal.xlsx");
    }
}
