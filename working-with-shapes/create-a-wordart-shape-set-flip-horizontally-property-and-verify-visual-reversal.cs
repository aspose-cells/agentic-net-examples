// Title: Create and Horizontally Flip a WordArt Shape with Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to add a WordArt shape to a worksheet using Aspose.Cells, verify the shape type, set the IsFlippedHorizontally property to true, and save the workbook so the mirrored text is visible in Excel.
// Keywords: Aspose.Cells C# | add WordArt shape | IsFlippedHorizontally | horizontal flip | PresetWordArtStyle | Excel shape manipulation | .NET Excel API
// Common Searches: Aspose.Cells flip WordArt horizontally | C# add WordArt to Excel worksheet | IsFlippedHorizontally property usage | How to mirror WordArt in Aspose.Cells | Save Excel file with flipped WordArt
// Developer Intent: Add a WordArt object to a worksheet, flip it horizontally, and verify the change.
// Use Cases: Create mirrored decorative text for report headers. | Support right‑to‑left language layouts by programmatically reversing WordArt. | Automate visual checks of shape orientation before exporting Excel files.
// AI Prompts: Write C# code that uses Aspose.Cells to insert a WordArt shape, set IsFlippedHorizontally to true, and save the workbook. | Explain the visual effect of the IsFlippedHorizontally property on WordArt in an Excel file. | Generate a reusable method that receives a Shape, confirms it is WordArt, and toggles its horizontal flip based on a boolean argument.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtFlipDemo
{
    // This example demonstrates how to add a WordArt shape to a worksheet using Aspose.Cells, verify the shape type, set the IsFlippedHorizontally property to true, and save the workbook so the mirrored text is visible in Excel.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a WordArt shape with a preset style
            // Parameters: style, text, topRow, top, leftColumn, left, height, width
            Shape wordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1, // preset style
                "Flip Me!",                     // text
                2, 0,                           // top row and vertical offset (pixels)
                2, 0,                           // left column and horizontal offset (pixels)
                100, 300);                      // height and width (pixels)

            // Verify that the shape is WordArt
            if (wordArt.IsWordArt)
            {
                // Flip the shape horizontally
                wordArt.IsFlippedHorizontally = true;

                // Output the flip status to the console for verification
                Console.WriteLine("IsFlippedHorizontally: " + wordArt.IsFlippedHorizontally);
            }
            else
            {
                Console.WriteLine("The created shape is not a WordArt object.");
            }

            // Save the workbook to visualize the flipped WordArt
            workbook.Save("WordArtFlippedHorizontally.xlsx");
        }
    }
}
