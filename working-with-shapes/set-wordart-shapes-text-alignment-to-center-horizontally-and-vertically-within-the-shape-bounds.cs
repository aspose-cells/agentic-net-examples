// Title: Center WordArt Text Horizontally & Vertically in Excel using Aspose.Cells for .NET (C#)
// Description: Shows how to insert a WordArt shape into an Excel workbook with Aspose.Cells for .NET and set its TextHorizontalAlignment and TextVerticalAlignment properties to Center before saving the file.
// Keywords: Aspose.Cells | C# | WordArt | text alignment | horizontal center | vertical center | Excel shape | PresetWordArtStyle | TextHorizontalAlignment | TextVerticalAlignment
// Common Searches: Aspose.Cells center WordArt text | C# set WordArt horizontal alignment in Excel | how to vertically align WordArt with Aspose.Cells | center text inside WordArt shape .NET | Aspose.Cells WordArt alignment example
// Developer Intent: Center the text of a WordArt shape both horizontally and vertically.
// Use Cases: Create a title banner where the WordArt label stays perfectly centered regardless of column width. | Generate dashboard cards with WordArt captions that remain aligned in the middle of each shape. | Build a reusable Excel template where WordArt captions automatically stay centered after shape resizing.
// AI Prompts: Provide C# code to align WordArt text to the left and top using Aspose.Cells. | Show how to resize a WordArt shape while preserving its centered text alignment. | Explain how to apply different PresetWordArtStyle values without affecting text centering.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtAlignment
{
    // Shows how to insert a WordArt shape into an Excel workbook with Aspose.Cells for .NET and set its TextHorizontalAlignment and TextVerticalAlignment properties to Center before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a WordArt shape
            // Parameters: style, text, topRow, top, leftColumn, left, height, width
            Shape wordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Centered WordArt",
                2,      // topRow
                10,     // top (pixels)
                2,      // leftColumn
                10,     // left (pixels)
                100,    // height (pixels)
                300);   // width (pixels)

            // Set horizontal text alignment to Center
            wordArt.TextHorizontalAlignment = TextAlignmentType.Center;

            // Set vertical text alignment to Center
            wordArt.TextVerticalAlignment = TextAlignmentType.Center;

            // Save the workbook
            workbook.Save("WordArtCentered.xlsx");
        }
    }
}
