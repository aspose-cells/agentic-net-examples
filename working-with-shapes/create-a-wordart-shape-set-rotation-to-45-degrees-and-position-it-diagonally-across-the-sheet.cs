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

        // Access the shapes collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add a WordArt shape with a preset style
        // Parameters: style, text, topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        Shape wordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset style
            "Diagonal WordArt",               // text
            0,    // topRow (row index)
            0,    // top (pixel offset from the topRow)
            0,    // leftColumn (column index)
            0,    // left (pixel offset from the leftColumn)
            500,  // height in pixels
            500   // width in pixels
        );

        // Rotate the WordArt shape by 45 degrees
        wordArt.RotationAngle = 45;

        // Save the workbook to a file
        workbook.Save("WordArtDiagonal.xlsx");
    }
}