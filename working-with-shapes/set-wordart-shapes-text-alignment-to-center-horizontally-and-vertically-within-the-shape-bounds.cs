using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtAlignment
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a WordArt shape
            // Parameters: style, text, topRow, top (pixels), leftColumn, left (pixels), height (pixels), width (pixels)
            Shape wordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Centered WordArt",
                2,      // topRow
                10,     // top offset in pixels
                2,      // leftColumn
                10,     // left offset in pixels
                100,    // height in pixels
                300     // width in pixels
            );

            // Set horizontal and vertical text alignment to center within the shape bounds
            wordArt.TextHorizontalAlignment = TextAlignmentType.Center;
            wordArt.TextVerticalAlignment = TextAlignmentType.Center;

            // Save the workbook
            workbook.Save("WordArtCenteredAlignment.xlsx");
        }
    }
}