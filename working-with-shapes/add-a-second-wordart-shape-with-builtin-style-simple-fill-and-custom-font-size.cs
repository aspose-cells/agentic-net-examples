// Title: Add a second WordArt shape with Simple Fill and custom font size using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, adds two WordArt shapes to the first worksheet, applies the Simple Fill preset (WordArtStyle4) to the second shape, sets its font size to 24 pt, and saves the file as SecondWordArt.xlsx.
// Keywords: Aspose.Cells C# WordArt | Add WordArt shape | Simple Fill WordArt | WordArtStyle4 | Set WordArt font size | TextEffect FontSize | multiple WordArt shapes | Excel shape collection | Aspose.Cells shape API
// Common Searches: how to add a second WordArt shape in Aspose.Cells | Aspose.Cells simple fill WordArt example | C# set WordArt font size Aspose.Cells | add multiple WordArt objects to an Excel worksheet | WordArtStyle4 Aspose.Cells C#
// Developer Intent: Insert a second WordArt shape with a Simple Fill preset and change its font size programmatically.
// Use Cases: Design a main title and a subtitle on the same sheet using different WordArt styles. | Create a report header with a plain‑fill WordArt shape and enlarge its text for emphasis. | Automate branding by adding several WordArt shapes with distinct presets and custom sizes in an Excel file.
// AI Prompts: Generate C# code that adds a WordArt shape using PresetWordArtStyle.WordArtStyle4, sets its text, and changes TextEffect.FontSize to a given value. | Explain how to verify a shape is WordArt with IsWordArt and modify its TextEffect properties in Aspose.Cells. | Show an example of adding multiple WordArt shapes with different presets and custom font sizes to a worksheet using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtExample
{
    // Creates a new workbook, adds two WordArt shapes to the first worksheet, applies the Simple Fill preset (WordArtStyle4) to the second shape, sets its font size to 24 pt, and saves the file as SecondWordArt.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the shape collection of the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add the first WordArt shape (any preset style)
            Shape firstWordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1, // preset style
                "First WordArt",                  // text
                1, 0,                             // top row, top offset (pixels)
                1, 0,                             // left column, left offset (pixels)
                100, 400);                        // height, width (pixels)

            // Add the second WordArt shape with a simple fill style (using a preset that has a plain fill)
            Shape secondWordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle4, // simple fill (white) preset
                "Second WordArt",                 // text
                5, 0,                             // top row, top offset
                5, 0,                             // left column, left offset
                100, 400);                        // height, width

            // Ensure the shape is WordArt and set a custom font size
            if (secondWordArt.IsWordArt)
            {
                // TextEffect provides access to WordArt formatting
                secondWordArt.TextEffect.FontSize = 24; // custom font size in points
            }

            // Save the workbook to a file
            workbook.Save("SecondWordArt.xlsx");
        }
    }
}
