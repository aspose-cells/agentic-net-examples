// Title: C# – Add WordArt to a Cell Range and Anchor It to Move & Resize with Cells using Aspose.Cells
// Description: Demonstrates how to create a workbook, insert a WordArt shape with a preset style, anchor it to a specific range (e.g., B2:D5) using MoveToRange, set its Placement to MoveAndSize so it moves and resizes with the cells, and save the file as WordArtAnchored.xlsx.
// Keywords: Aspose.Cells WordArt C# | add WordArt shape Aspose.Cells | anchor shape to cell range | MoveToRange Aspose.Cells | PlacementType.MoveAndSize | Excel shape placement .NET | WordArt MoveAndSize example | Aspose.Cells shape anchoring
// Common Searches: How to insert WordArt into a specific range with Aspose.Cells .NET | Aspose.Cells MoveToRange shape example | Set WordArt to move and resize with cells in Excel using Aspose.Cells | Aspose.Cells shape placement types MoveAndSize | C# code for anchoring WordArt to cells
// Developer Intent: Insert a WordArt shape, bind it to a defined cell range, and configure it to move and resize together with the cells.
// Use Cases: Create a branded header that stays aligned with a table when rows are added or removed. | Add decorative WordArt to a report that automatically adjusts to layout changes. | Place a dynamic title that scales with a chart area as the chart expands.
// AI Prompts: Generate C# code with Aspose.Cells to add a WordArt shape to range C3:E6 and set its placement to MoveAndSize. | Explain the effect of MoveToRange and Placement properties on shape behavior in Aspose.Cells workbooks. | Provide a step‑by‑step guide to anchor multiple WordArt shapes to different cell ranges in a single workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtExample
{
    // Demonstrates how to create a workbook, insert a WordArt shape with a preset style, anchor it to a specific range (e.g., B2:D5) using MoveToRange, set its Placement to MoveAndSize so it moves and resizes with the cells, and save the file as WordArtAnchored.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the shape collection of the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add a WordArt shape (preset style, text, position, size)
            // Parameters: style, text, topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
            Shape wordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Aspose.Cells WordArt",
                0,      // topRow (upper left row index)
                0,      // top pixel offset
                0,      // leftColumn (upper left column index)
                0,      // left pixel offset
                100,    // height in pixels
                300);   // width in pixels

            // Anchor the WordArt to a specific cell range (e.g., B2:D5)
            // MoveToRange(startRow, startColumn, endRow, endColumn)
            wordArt.MoveToRange(1, 1, 4, 3); // B2 (row 1, col 1) to D5 (row 4, col 3)

            // Set the placement so the shape moves and resizes with the cells
            wordArt.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            workbook.Save("WordArtAnchored.xlsx");
        }
    }
}
