using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add the second WordArt shape.
        // Using PresetWordArtStyle.WordArtStyle4 as a simple fill style.
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape secondWordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle4,
            "Second WordArt",
            2,      // topRow
            0,      // top (pixel offset)
            2,      // leftColumn
            0,      // left (pixel offset)
            100,    // height (pixels)
            300);   // width (pixels)

        // Set a custom font size for the WordArt text
        if (secondWordArt.IsWordArt)
        {
            secondWordArt.TextEffect.FontSize = 24; // font size in points
        }

        // Save the workbook
        workbook.Save("SecondWordArt.xlsx");
    }
}