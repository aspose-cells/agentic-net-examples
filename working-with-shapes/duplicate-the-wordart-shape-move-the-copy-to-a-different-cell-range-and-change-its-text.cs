using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class DuplicateWordArtExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the shapes collection of the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add an original WordArt shape
            // Parameters: style, text, upperLeftRow, upperLeftColumn, rowOffset, columnOffset, height, width
            Shape originalWordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Original WordArt",
                2,   // upperLeftRow (zero‑based)
                2,   // upperLeftColumn (zero‑based)
                0,   // row offset in pixels
                0,   // column offset in pixels
                100, // height in pixels
                200  // width in pixels
            );

            // Duplicate the WordArt shape with the same style and size at a different location
            Shape duplicatedWordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                originalWordArt.TextEffect.Text, // copy original text
                5,   // new upperLeftRow
                5,   // new upperLeftColumn
                0,   // row offset
                0,   // column offset
                100, // height
                200  // width
            );

            // Change the text of the duplicated WordArt
            duplicatedWordArt.TextEffect.Text = "Duplicated WordArt";

            // Optionally move the duplicated shape to a specific cell range (e.g., B6:D8)
            // MoveToRange(startRow, startColumn, endRow, endColumn) – rows/columns are zero‑based
            duplicatedWordArt.MoveToRange(5, 5, 7, 7);

            // Save the workbook
            workbook.Save("DuplicateWordArt.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}