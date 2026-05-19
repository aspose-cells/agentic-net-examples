using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AdjustWordArtToWorksheet
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with initial dummy size
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset style
            "Full Sheet WordArt",             // text
            0,   // topRow
            0,   // top (pixel offset)
            0,   // leftColumn
            0,   // left (pixel offset)
            0,   // height (pixel)
            0);  // width (pixel)

        // Set the shape to start at the top‑left corner of the worksheet
        wordArt.Top = 0;   // vertical offset in pixels
        wordArt.Left = 0;  // horizontal offset in pixels

        // Approximate the worksheet size in pixels.
        // One column width is roughly 64 pixels and one row height is roughly 20 pixels.
        int totalColumns = worksheet.Cells.MaxColumn + 1; // MaxColumn is zero‑based
        int totalRows = worksheet.Cells.MaxRow + 1;       // MaxRow is zero‑based

        // Expand the shape to cover the whole worksheet
        wordArt.Width = totalColumns * 64;   // width in pixels
        wordArt.Height = totalRows * 20;     // height in pixels

        // Optionally, fit the text to the new size
        wordArt.FitToTextSize();

        // Save the workbook
        workbook.Save("WordArtFullWorksheet.xlsx");
    }
}