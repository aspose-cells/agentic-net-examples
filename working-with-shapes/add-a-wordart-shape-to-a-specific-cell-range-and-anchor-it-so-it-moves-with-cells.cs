// Title: Add WordArt to a Cell Range and Anchor It with MoveAndSize using Aspose.Cells for .NET
// Description: Step‑by‑step guide to insert a WordArt shape, anchor it to a specific range (e.g., B2:D5) and set PlacementType.MoveAndSize so the shape moves and resizes with the cells in an Excel workbook via Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# WordArt | AddWordArt shape | MoveToRange | PlacementType.MoveAndSize | anchor WordArt to cells | Excel shape anchoring | Aspose.Cells shape example | C# Excel WordArt | Aspose.Cells shape placement
// Common Searches: Aspose.Cells add WordArt to Excel | How to anchor WordArt to a range in Aspose.Cells | C# MoveAndSize placement for shapes | Set WordArt to move with cells Aspose | Aspose.Cells shape MoveToRange example
// Developer Intent: Insert a WordArt shape, bind it to a defined cell range, and configure it to move and resize with the worksheet cells.
// Use Cases: Create a branded header that stays aligned with a table when rows are added or removed. | Add a dynamic chart title as WordArt that automatically adjusts its position and size with the chart area. | Place a watermark WordArt that expands or contracts with the printable area of the worksheet.
// AI Prompts: Generate C# code with Aspose.Cells that adds WordArt to range C3:E6, sets the fill color to LightGreen, and applies PlacementType.MoveAndSize. | Explain how PlacementType.MoveAndSize influences a WordArt shape when rows or columns are inserted or deleted in an Excel sheet. | Provide a reusable C# method that creates WordArt with custom text, style, and color, anchors it to any given cell range, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtExample
{
    // Step‑by‑step guide to insert a WordArt shape, anchor it to a specific range (e.g., B2:D5) and set PlacementType.MoveAndSize so the shape moves and resizes with the cells in an Excel workbook via Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the shape collection of the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add a WordArt shape (preset style, text, initial position and size in pixels)
            Shape wordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1, // preset style
                "Hello Aspose!",                 // text
                0,   // topRow (initial row index)
                0,   // top offset in pixels
                0,   // leftColumn (initial column index)
                0,   // left offset in pixels
                100, // height in pixels
                300  // width in pixels
            );

            // Anchor the WordArt to a specific cell range (e.g., B2:D5)
            // This makes the shape move and resize with the range.
            int startRow = 1;    // B2 -> row index 1 (zero‑based)
            int startColumn = 1; // B2 -> column index 1
            int endRow = 4;      // D5 -> row index 4
            int endColumn = 3;   // D5 -> column index 3
            wordArt.MoveToRange(startRow, startColumn, endRow, endColumn);

            // Set the placement so the shape moves and sizes with the cells
            wordArt.Placement = PlacementType.MoveAndSize;

            // Optional: customize appearance (e.g., fill color)
            wordArt.FillFormat.ForeColor = System.Drawing.Color.LightBlue;

            // Save the workbook
            workbook.Save("WordArtAnchored.xlsx");
        }
    }
}
