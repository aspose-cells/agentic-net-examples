// Title: Insert WordArt into a cell range and lock it to move with cells using Aspose.Cells for .NET
// Common Searches: how to add WordArt to a specific range in an Excel file with Aspose.Cells C# | Aspose.Cells set WordArt shape to move with cells | resize WordArt shape to match Excel range dimensions using Aspose.Cells | calculate WordArt width and height from column widths and row heights in C# | place WordArt at B2 and anchor it to cells with Aspose.Cells .NET
// Tags: add WordArt shape Aspose.Cells | WordArt placement Move Aspose.Cells | anchor WordArt to cell range | adjust WordArt size to range dimensions | derive shape dimensions from column and row metrics C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a WordArt shape anchored to the upper‑left cell of a defined range (B2:D5), sets its placement to Move so it follows cell shifts, computes the shape's width and height from column widths and row heights, applies these dimensions, and saves the file as WordArtExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the target cell range (e.g., B2:D5)
            int startRow = 1;      // Row 2 (zero‑based)
            int startColumn = 1;   // Column B (zero‑based)
            int endRow = 4;        // Row 5
            int endColumn = 3;     // Column D

            // Add a WordArt shape anchored to the upper‑left cell of the range
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1, // preset WordArt style
                "Aspose WordArt",                 // text
                startRow, startColumn,           // upper‑left cell
                0, 0,                             // offsets within the cell
                100, 300);                        // height, width in points

            // Make the shape move with cells (it will stay attached when rows/columns shift)
            wordArt.Placement = PlacementType.Move;

            // Optionally resize the shape to fit the target range
            // Approximate column width in points (1 character ≈ 7 points)
            double width = (sheet.Cells.GetColumnWidth(endColumn) - sheet.Cells.GetColumnWidth(startColumn)) * 7;
            // Row height is already in points
            double height = sheet.Cells.GetRowHeight(endRow) - sheet.Cells.GetRowHeight(startRow);

            // Shape.Width and Height expect integer values (points), so cast explicitly
            wordArt.Width = (int)width;
            wordArt.Height = (int)height;

            // Save the workbook
            workbook.Save("WordArtExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
