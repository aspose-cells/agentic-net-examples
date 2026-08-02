// Title: Add a rectangle shape to a specific cell with Aspose.Cells for .NET
// Description: Shows how to create a workbook, locate cell C5 (row 5, column C), and insert a rectangle whose top‑left corner aligns with that cell. The sample sets pixel‑based width and height, applies a solid fill, border style, and label text, then saves the file as RectangleAtCell.xlsx.
// Keywords: Aspose.Cells rectangle shape | C# add shape to cell | shape anchoring Aspose.Cells | Excel drawing objects .NET | set shape fill color | shape dimensions pixels | AddRectangle method | Aspose.Cells worksheet graphics
// Common Searches: Aspose.Cells add rectangle to cell | How to anchor a shape to a cell in C# | Set shape size and style with Aspose.Cells | Place drawing objects at exact cell coordinates | C# Aspose.Cells shape positioning example
// Developer Intent: Insert a rectangle drawing object positioned at a given worksheet cell.
// Use Cases: Highlight a data block by placing a colored rectangle over the top‑left cell of the range. | Design a template with visual sections by anchoring rectangles to predefined cells. | Generate reports that add callout boxes to specific cells, customizing size, fill, and border.
// AI Prompts: Write C# code using Aspose.Cells to add a rectangle at cell D10 with a width of 200 px, height of 80 px, a red border, and the text "Important". | Show how to iterate over a list of cell addresses and attach a rectangle shape to each one with Aspose.Cells for .NET. | Explain the conversion from Excel cell references (e.g., "B3") to the topRow, leftColumn, top, and left parameters required by the AddRectangle method.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeDemo
{
    // Shows how to create a workbook, locate cell C5 (row 5, column C), and insert a rectangle whose top‑left corner aligns with that cell. The sample sets pixel‑based width and height, applies a solid fill, border style, and label text, then saves the file as RectangleAtCell.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the target cell (e.g., "C5")
            // Row and column indexes are zero‑based
            int targetRow = 4;    // Row 5 in Excel (C5)
            int targetColumn = 2; // Column C in Excel

            // Define shape dimensions (in pixels)
            int shapeHeight = 100;
            int shapeWidth = 150;

            // Add a rectangle shape anchored to the target cell
            // top and left offsets are set to 0 (shape starts exactly at the cell's top‑left corner)
            RectangleShape rectangle = worksheet.Shapes.AddRectangle(
                topRow: targetRow,
                top: 0,
                leftColumn: targetColumn,
                left: 0,
                height: shapeHeight,
                width: shapeWidth);

            // Optional: set some visual properties
            rectangle.Fill.FillType = FillType.Solid;
            rectangle.Fill.SolidFill.Color = System.Drawing.Color.LightGreen;
            rectangle.Line.DashStyle = MsoLineDashStyle.Solid;
            rectangle.Line.Weight = 2;
            rectangle.Text = "Sample Rectangle";

            // Save the workbook
            workbook.Save("RectangleAtCell.xlsx");
        }
    }
}
