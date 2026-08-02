// Title: Center WordArt Text Horizontally and Vertically with Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, inserts a WordArt shape using a preset style, and centers the text inside the shape by setting the TextHorizontalAlignment and TextVerticalAlignment properties to Center. The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells WordArt alignment | C# center WordArt text | TextHorizontalAlignment Aspose.Cells | TextVerticalAlignment Aspose.Cells | AddWordArt C# | Excel WordArt center text | Aspose.Cells shape formatting
// Common Searches: Aspose.Cells center WordArt text horizontally | how to vertically align WordArt text in .NET | C# set WordArt text alignment Aspose.Cells | center text inside WordArt shape Excel | Aspose.Cells WordArt alignment example
// Developer Intent: Align WordArt text to the middle of the shape both horizontally and vertically using Aspose.Cells for .NET.
// Use Cases: Generate a report title with WordArt that remains centered regardless of shape resizing. | Design a dashboard label where the WordArt caption must stay perfectly centered within its bounds.
// AI Prompts: Write C# code to align WordArt text to the left and top using Aspose.Cells. | Show how to retrieve the current TextHorizontalAlignment and TextVerticalAlignment values from a WordArt shape. | Create an example that applies different preset WordArt styles while keeping the text centered.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtAlignment
{
    // This example creates a new workbook, inserts a WordArt shape using a preset style, and centers the text inside the shape by setting the TextHorizontalAlignment and TextVerticalAlignment properties to Center. The workbook is then saved as an Excel file.
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

            // Set horizontal and vertical text alignment to Center
            wordArt.TextHorizontalAlignment = TextAlignmentType.Center;
            wordArt.TextVerticalAlignment = TextAlignmentType.Center;

            // Save the workbook
            workbook.Save("WordArtCenteredAlignment.xlsx");
        }
    }
}
