// Title: Aspose.Cells for .NET – C# code to make a WordArt shape fill the whole worksheet
// Description: Demonstrates how to add a WordArt shape to a workbook, compute the worksheet's used range with MaxDataRow/MaxDataColumn, and use MoveToRange to set the shape's top, left, width, and height so it covers every populated cell before saving the file.
// Keywords: Aspose.Cells WordArt fill worksheet | C# WordArt shape resize Aspose | MoveToRange WordArt .NET | set shape dimensions worksheet Aspose.Cells | full‑sheet WordArt example | Aspose.Cells used range sizing
// Common Searches: Aspose.Cells make WordArt cover entire sheet | C# resize WordArt to used range | MoveToRange example for shapes Aspose.Cells | how to set WordArt size programmatically .NET
// Developer Intent: Programmatically adjust a WordArt shape so its bounds match the worksheet's used area.
// Use Cases: Create a full‑sheet header or banner that automatically expands with added data. | Apply a watermark WordArt that always spans all rows and columns containing values. | Refresh an existing WordArt after data import to keep it aligned with the current data range.
// AI Prompts: Generate C# code using Aspose.Cells that positions a WordArt shape to cover the entire used range of a worksheet. | Show how to handle an empty worksheet when resizing a WordArt shape with MoveToRange in Aspose.Cells. | Explain how MaxDataRow and MaxDataColumn are used to size shapes in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a WordArt shape to a workbook, compute the worksheet's used range with MaxDataRow/MaxDataColumn, and use MoveToRange to set the shape's top, left, width, and height so it covers every populated cell before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with initial zero size and position
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset style
            "Full Sheet",                     // text
            0,  // topRow (row index)
            0,  // top offset in pixels
            0,  // leftColumn (column index)
            0,  // left offset in pixels
            0,  // height in pixels
            0   // width in pixels
        );

        // Determine the used range of the worksheet
        int maxRow = worksheet.Cells.MaxDataRow;       // last used row index
        int maxCol = worksheet.Cells.MaxDataColumn;    // last used column index

        // Resize and reposition the WordArt to cover the entire used range
        // MoveToRange sets the shape to span from (0,0) to (maxRow, maxCol)
        wordArt.MoveToRange(0, 0, maxRow, maxCol);

        // Save the workbook
        workbook.Save("WordArtFullSheet.xlsx");
    }
}
