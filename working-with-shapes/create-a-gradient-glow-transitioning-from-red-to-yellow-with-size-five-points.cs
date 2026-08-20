// Title: Apply a Red‑to‑Yellow Gradient Fill and 5‑Point Glow to a Rectangle Shape with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a rectangle shape, sets a horizontal two‑color gradient from red to yellow, adds a 5‑point red glow, and saves the file as GradientGlowDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | gradient fill | shape glow | rectangle shape | Excel workbook | two‑color gradient | horizontal gradient | glow size points | red glow | .NET
// Common Searches: Aspose.Cells add gradient fill to shape | C# set shape glow size Aspose.Cells | rectangle shape red to yellow gradient Excel | how to apply glow effect to shape in Aspose.Cells | gradient and glow example Aspose.Cells .NET
// Developer Intent: Add a rectangle shape to an Excel file, fill it with a red‑to‑yellow gradient, and apply a 5‑point red glow using Aspose.Cells for .NET.
// Use Cases: Enhance KPI dashboards by highlighting key cells with gradient‑filled shapes that have a subtle glow. | Generate marketing or sales flyers in Excel where colored shapes draw attention to promotional messages. | Automate report templates that use styled rectangles to separate sections and improve visual hierarchy.
// AI Prompts: Generate C# code that creates an oval shape with a blue‑to‑green vertical gradient and an 8‑point green glow using Aspose.Cells. | Explain how to change a shape's glow color dynamically based on a cell value in an Aspose.Cells workbook. | Provide steps to export an Excel workbook containing gradient‑glow shapes to PDF while preserving the visual effects.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGradientGlowDemo
{
    // Creates a new workbook, inserts a rectangle shape, sets a horizontal two‑color gradient from red to yellow, adds a 5‑point red glow, and saves the file as GradientGlowDemo.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to demonstrate the effect
            // Parameters: drawing type, upper left row, upper left column, top, left, width, height
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);

            // Set the shape's fill to a two‑color gradient (red → yellow)
            shape.Fill.FillType = FillType.Gradient;
            GradientFill gradientFill = shape.Fill.GradientFill;
            gradientFill.SetTwoColorGradient(
                Color.Red,          // start color
                Color.Yellow,       // end color
                GradientStyleType.Horizontal, // gradient direction
                1);                 // variant (default)

            // Configure the glow effect: size = 5 points, color = red (can be any color)
            shape.Glow.Size = 5;                     // radius in points
            shape.Glow.Color = workbook.CreateCellsColor();
            shape.Glow.Color.Color = Color.Red;     // glow color

            // Save the workbook
            workbook.Save("GradientGlowDemo.xlsx");
        }
    }
}
