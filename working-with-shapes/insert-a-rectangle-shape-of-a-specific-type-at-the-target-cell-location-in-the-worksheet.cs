// Title: Add a rectangle shape to a target cell with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, selects the first worksheet, defines cell C5 as the anchor point, sets rectangle dimensions in pixels, adds a rectangle shape with zero offsets so it aligns to the cell's top‑left corner, applies fill and line colors, adds a label, and saves the file as RectangleAtCell.xlsx.
// Keywords: Aspose.Cells rectangle shape C# | add shape to cell Aspose.Cells | shape positioning Aspose.Cells | Aspose.Cells rectangle fill color | Aspose.Cells shape example
// Common Searches: how to add a rectangle shape to a specific cell using Aspose.Cells | Aspose.Cells place shape at cell C5 with no offset | set rectangle size and style in Aspose.Cells workbook | anchor drawing shape to worksheet cell C# | Aspose.Cells shape alignment tutorial
// Developer Intent: Insert and style a rectangle shape that is anchored to a designated worksheet cell.
// Use Cases: Highlight a key metric by surrounding its cell with a colored rectangle in an automated report. | Create a labeled box next to a data entry cell for visual guidance in generated spreadsheets. | Design a template where shapes are pre‑positioned to align with form fields for bulk document creation.
// AI Prompts: Generate C# code with Aspose.Cells to insert a rectangle at cell D10, specifying custom width, height, fill, and line colors. | Show how to anchor multiple rectangle shapes to different cells in a worksheet using Aspose.Cells for .NET. | Provide an example that updates the text of an existing rectangle shape positioned at a given cell.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeExample
{
    // Creates a new workbook, selects the first worksheet, defines cell C5 as the anchor point, sets rectangle dimensions in pixels, adds a rectangle shape with zero offsets so it aligns to the cell's top‑left corner, applies fill and line colors, adds a label, and saves the file as RectangleAtCell.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the target cell (e.g., C5). 
            // Row and column indices are zero‑based: C -> 2, 5 -> 4
            int targetRow = 4;      // Row index for cell C5
            int targetColumn = 2;   // Column index for cell C5

            // Define shape size in pixels
            int shapeHeight = 100;  // Height of the rectangle
            int shapeWidth = 200;   // Width of the rectangle

            // Add a rectangle shape anchored to the target cell.
            // Offsets (top, left) are set to 0 to align the shape with the cell's top‑left corner.
            RectangleShape rectangle = worksheet.Shapes.AddRectangle(
                topRow: targetRow,
                top: 0,
                leftColumn: targetColumn,
                left: 0,
                height: shapeHeight,
                width: shapeWidth);

            // Optional: set some visual properties
            rectangle.Fill.SolidFill.Color = System.Drawing.Color.LightBlue;
            rectangle.Line.SolidFill.Color = System.Drawing.Color.DarkBlue;
            rectangle.Line.Weight = 2;
            rectangle.Text = "Sample Rectangle";

            // Save the workbook
            workbook.Save("RectangleAtCell.xlsx");
        }
    }
}
