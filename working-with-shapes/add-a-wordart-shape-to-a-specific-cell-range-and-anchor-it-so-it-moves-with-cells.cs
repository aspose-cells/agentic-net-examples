// Title: Insert WordArt into a worksheet range and anchor it to move and resize with cells using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add WordArt shape to a specific cell range | anchor WordArt shape to cells so it moves and resizes with them in Aspose.Cells | set shape placement to MoveAndSize in Aspose.Cells .NET example | position WordArt between B2 and D5 using Aspose.Cells C# | use UpperLeftRow UpperLeftColumn LowerRightRow LowerRightColumn to bind shape in Aspose.Cells
// Tags: WordArt shape insertion Aspose.Cells | anchor shape to cell range in Excel workbook | configure shape to move and resize with cells | set shape coordinates UpperLeftRow UpperLeftColumn LowerRightRow LowerRightColumn | customize WordArt font and color using Aspose.Cells C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a WordArt shape to the B2:D5 range, anchors it by setting UpperLeftRow/Column and LowerRightRow/Column, sets the Placement to MoveAndSize so the shape moves and resizes with the cells, customizes the font appearance, and saves the file as WordArtAnchored.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell range that will anchor the WordArt (e.g., B2:D5)
            // Row and column indices are zero‑based.
            int startRow = 1;      // B2 row
            int startColumn = 1;   // B2 column
            int endRow = 4;        // D5 row
            int endColumn = 3;     // D5 column

            // Define the visual size of the WordArt in pixels.
            // These values can be adjusted as needed.
            int height = 100;
            int width = 300;

            // Add a WordArt shape to the worksheet.
            // Parameters: style, text, topRow, top offset, leftColumn, left offset, height, width
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Aspose WordArt",
                startRow,      // topRow
                0,             // top offset (pixels)
                startColumn,   // leftColumn
                0,             // left offset (pixels)
                height,
                width);

            // Anchor the shape to the desired cell range.
            wordArt.UpperLeftRow = startRow;
            wordArt.UpperLeftColumn = startColumn;
            wordArt.LowerRightRow = endRow;
            wordArt.LowerRightColumn = endColumn;

            // Set the placement so the shape moves and resizes with the cells.
            wordArt.Placement = PlacementType.MoveAndSize;

            // Optional: customize appearance.
            wordArt.Font.Name = "Arial";
            wordArt.Font.Size = 24;
            wordArt.Font.Color = Color.Blue;

            // Save the workbook.
            const string outputPath = "WordArtAnchored.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
