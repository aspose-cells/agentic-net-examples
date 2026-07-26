// Title: Resize a Line Shape Across B2:E2 and Set Thickness with Aspose.Cells for .NET
// Description: Demonstrates adding a LineShape to a workbook, positioning it from cell B2 to E2, adjusting its weight to 2 points, and saving the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells line shape | resize line shape | line shape thickness | position line shape cells | C# Aspose.Cells line | Excel shape line weight | horizontal line B2 E2 | Aspose.Cells drawing API | set line weight points | add line shape .NET
// Common Searches: Aspose.Cells how to resize line shape | set line thickness Aspose.Cells C# | position line shape across cells B2 E2 | add horizontal line in Excel using Aspose.Cells | Aspose.Cells line shape coordinates | change line weight in Aspose.Cells .NET
// Developer Intent: Resize a line shape to cover cells B2‑E2 and change its weight.
// Use Cases: Add a separator line across a header row in automated reports | Visually connect two columns in a generated spreadsheet | Emphasize a data range with a bold line in an Excel export | Create a custom chart axis line programmatically
// AI Prompts: Write C# code that creates a LineShape from C3 to G3 with a 3‑point weight using Aspose.Cells. | Show how to calculate line shape coordinates from dynamic row/column indices in Aspose.Cells for .NET. | Explain converting pixel dimensions to point weight for line shapes in Aspose.Cells. | Provide a method to resize any shape to a given cell range and set its line style with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates adding a LineShape to a workbook, positioning it from cell B2 to E2, adjusting its weight to 2 points, and saving the file using Aspose.Cells for .NET.
class ResizeLineExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a line shape (initial size will be adjusted later)
            LineShape lineShape = worksheet.Shapes.AddLine(0, 0, 0, 0, 0, 0);

            // Position the line to span cells B2 (row 1, column 1) through E2 (row 1, column 4)
            lineShape.UpperLeftRow = 1;      // B2 row index (zero‑based)
            lineShape.UpperLeftColumn = 1;   // B column index
            lineShape.LowerRightRow = 1;     // Same row for a horizontal line
            lineShape.LowerRightColumn = 4;  // E column index

            // Adjust the line thickness (weight) in points
            lineShape.Line.Weight = 2.0f;    // 2 points thickness

            // Save the workbook
            workbook.Save("ResizedLine.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
