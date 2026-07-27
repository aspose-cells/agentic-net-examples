// Title: Flip a WordArt shape horizontally with Aspose.Cells for .NET
// Description: Demonstrates how to add a WordArt shape to a worksheet, set its IsFlippedHorizontally property to true, output the flag for verification, and save the workbook as an Excel file.
// Keywords: Aspose.Cells WordArt flip | C# flip WordArt horizontally | IsFlippedHorizontally property | Excel shape transformation Aspose | save workbook after shape edit
// Common Searches: how to flip WordArt horizontally using Aspose.Cells | Aspose.Cells C# set IsFlippedHorizontally | verify WordArt flip property in .NET | add WordArt shape to Excel with Aspose | save Excel file after modifying shape
// Developer Intent: Insert a WordArt object, mirror it on the horizontal axis, confirm the flip flag, and generate the resulting Excel file.
// Use Cases: Create mirrored text for branding in automated Excel reports. | Generate reversed logo graphics as WordArt for printable marketing materials. | Programmatically validate shape orientation before exporting to ensure design fidelity.
// AI Prompts: Show C# code to flip a WordArt shape vertically with Aspose.Cells. | Provide an example that toggles IsFlippedHorizontally based on a runtime condition. | Explain how to rotate, scale, or skew shapes using Aspose.Cells transformation properties.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a WordArt shape to a worksheet, set its IsFlippedHorizontally property to true, output the flag for verification, and save the workbook as an Excel file.
class WordArtFlipDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset style
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset style
            "Flip Me",                        // text
            2,    // top row index
            0,    // vertical offset (pixels)
            2,    // left column index
            0,    // horizontal offset (pixels)
            100,  // height (pixels)
            300   // width (pixels)
        );

        // Flip the WordArt horizontally
        wordArt.IsFlippedHorizontally = true;

        // Verify the flip property
        Console.WriteLine("IsFlippedHorizontally: " + wordArt.IsFlippedHorizontally);

        // Save the workbook to a file
        workbook.Save("WordArtFlipped.xlsx");
    }
}
