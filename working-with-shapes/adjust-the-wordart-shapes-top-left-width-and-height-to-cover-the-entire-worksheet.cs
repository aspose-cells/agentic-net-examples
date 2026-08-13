// Title: C# – Resize WordArt to Fill an Entire Worksheet with Aspose.Cells
// Description: Creates a workbook, adds a WordArt shape, calculates the worksheet's total pixel width and height, and sets the shape's Top, Left, Width, and Height so it covers the whole sheet before saving the file.
// Keywords: Aspose.Cells WordArt resize | C# set shape size Excel | fit shape to worksheet | calculate worksheet pixel dimensions | .NET Excel shape positioning
// Common Searches: Aspose.Cells resize WordArt to full sheet | C# get total worksheet width in pixels | set shape top left width height Aspose.Cells | cover entire worksheet with WordArt
// Developer Intent: Make a WordArt shape span the full width and height of a worksheet programmatically.
// Use Cases: Add a full‑sheet watermark or title using WordArt for report templates. | Automatically adjust any inserted shape to match dynamic worksheet sizes. | Apply consistent full‑sheet WordArt formatting across multiple worksheets in a workbook.
// AI Prompts: Write C# code using Aspose.Cells that resizes a WordArt shape to the worksheet's pixel dimensions, handling merged cells. | Explain how to efficiently compute a worksheet's total pixel width and height with Aspose.Cells APIs. | Show how to loop through all worksheets in a workbook and apply full‑sheet WordArt resizing to each sheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a WordArt shape, calculates the worksheet's total pixel width and height, and sets the shape's Top, Left, Width, and Height so it covers the whole sheet before saving the file.
class AdjustWordArtToWorksheet
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with temporary zero size at the top‑left corner
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset style
            "Full Worksheet",                // text
            0,  // topRow
            0,  // top (pixel offset)
            0,  // leftColumn
            0,  // left (pixel offset)
            0,  // height (pixel)
            0   // width (pixel)
        );

        // Calculate total width of all columns in pixels
        int totalWidth = 0;
        for (int col = 0; col <= worksheet.Cells.MaxColumn; col++)
        {
            totalWidth += worksheet.Cells.GetColumnWidthPixel(col);
        }

        // Calculate total height of all rows in pixels
        int totalHeight = 0;
        for (int row = 0; row <= worksheet.Cells.MaxRow; row++)
        {
            totalHeight += worksheet.Cells.GetRowHeightPixel(row);
        }

        // Adjust the WordArt to cover the entire worksheet
        wordArt.Top = 0;          // top edge at the worksheet top
        wordArt.Left = 0;         // left edge at the worksheet left
        wordArt.Width = totalWidth;
        wordArt.Height = totalHeight;

        // Save the workbook
        workbook.Save("WordArtFullWorksheet.xlsx");
    }
}
