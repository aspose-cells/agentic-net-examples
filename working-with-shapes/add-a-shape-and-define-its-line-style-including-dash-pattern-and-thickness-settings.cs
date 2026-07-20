// Title: Aspose.Cells for .NET – Add a Rectangle Shape and Configure Line Dash Style, Thickness, and Color
// Description: This example creates a new workbook, inserts a rectangle shape on the first worksheet, enables its border, and customizes the line by applying a DashDot pattern, a 3‑point weight, and a dark‑blue solid fill before saving the file as ShapeWithLineStyle.xlsx.
// Keywords: Aspose.Cells add shape .NET | shape line dash style | line thickness Aspose.Cells | shape border color C# | custom line format Excel | rectangle shape Aspose.Cells
// Common Searches: how to set dash style for a shape line in Aspose.Cells | Aspose.Cells change shape border thickness C# | add rectangle with custom line style Aspose.Cells | set shape line color and dash pattern .NET
// Developer Intent: Insert a shape into a worksheet and define its line dash pattern, weight, and color.
// Use Cases: Highlight a data range with a dash‑dot rectangle to draw attention in a report. | Create a legend box that uses a thick colored line for clear visual separation. | Add a callout shape with a custom dash pattern to annotate chart elements.
// AI Prompts: Write C# code using Aspose.Cells to add an ellipse shape with a dotted 2‑point red line. | Show how to modify an existing shape’s line to a solid 5‑point green border in Aspose.Cells. | Demonstrate applying different dash styles to multiple shapes on the same worksheet with Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLineDemo
{
    // This example creates a new workbook, inserts a rectangle shape on the first worksheet, enables its border, and customizes the line by applying a DashDot pattern, a 3‑point weight, and a dark‑blue solid fill before saving the file as ShapeWithLineStyle.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: topRow, top, leftColumn, left, height, width
            Shape shape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 120, 80);

            // Ensure the shape has a visible line border
            shape.HasLine = true;

            // Access the line format via the Shape.Line property
            LineFormat line = shape.Line;

            // Set dash pattern (e.g., DashDot) and line thickness (weight in points)
            line.DashStyle = MsoLineDashStyle.DashDot;
            line.Weight = 3.0; // thickness of 3 points

            // Optionally set line color for better visibility
            line.SolidFill.Color = Color.DarkBlue;

            // Save the workbook to a file
            workbook.Save("ShapeWithLineStyle.xlsx");
        }
    }
}
