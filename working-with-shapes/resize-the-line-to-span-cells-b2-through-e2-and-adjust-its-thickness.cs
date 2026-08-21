// Title: Aspose.Cells .NET – Resize a LineShape to span B2‑E2 and set its thickness
// Description: Creates a workbook, adds a LineShape, positions it from cell B2 to E2 by setting row/column indices and zero pixel offsets, adjusts the line weight (e.g., 2.5 pt), and saves the file as LineResized.xlsx.
// Keywords: Aspose.Cells | .NET | C# | LineShape | resize line shape | line thickness | line weight | position shape across cells | B2 to E2 | shape pixel offsets
// Common Searches: Aspose.Cells resize line shape across cells | Set line weight in Aspose.Cells .NET | Position LineShape from B2 to E2 using Aspose.Cells | How to remove pixel offsets for a shape in Aspose.Cells | C# line shape spanning multiple columns Aspose.Cells
// Developer Intent: Resize a LineShape to cover cells B2‑E2 and change its line weight.
// Use Cases: Add a horizontal separator line across a specific row in an automated report. | Emphasize a range by drawing a thick line over selected cells for visual cues. | Match corporate branding by applying custom line thickness to spreadsheet graphics.
// AI Prompts: Generate C# code that places a LineShape from B2 to E2 with zero pixel offsets and sets its weight to 3 pt using Aspose.Cells. | Show how to programmatically adjust a line's thickness based on a variable while keeping it anchored across cells B2‑E2. | Explain the steps to resize any shape to span a given cell range and modify its line weight in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a LineShape, positions it from cell B2 to E2 by setting row/column indices and zero pixel offsets, adjusts the line weight (e.g., 2.5 pt), and saves the file as LineResized.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a line shape (initial size will be adjusted later)
            // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height (px), width (px)
            LineShape line = sheet.Shapes.AddLine(1, 0, 1, 0, 0, 0);

            // Position the line to span cells B2 (column 1) through E2 (column 4) on the same row (row 1)
            line.UpperLeftRow = 1;      // Row index for B2 (zero‑based)
            line.UpperLeftColumn = 1;   // Column index for B
            line.LowerRightRow = 1;     // Same row for the end point
            line.LowerRightColumn = 4;  // Column index for E

            // Ensure there are no extra pixel offsets
            line.UpperDeltaX = 0;
            line.UpperDeltaY = 0;
            line.LowerDeltaX = 0;
            line.LowerDeltaY = 0;

            // Adjust the line thickness (weight) in points
            line.Line.Weight = 2.5f; // Example thickness of 2.5 pt

            // Save the workbook with the resized line
            workbook.Save("LineResized.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
