// Title: Create a rectangle with a red‑to‑yellow gradient fill and 5‑point glow using Aspose.Cells for .NET (C#)
// Description: Shows how to add a rectangle shape to the first worksheet, apply a horizontal red‑to‑yellow two‑color gradient fill, set a 5‑point yellow glow, and save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | gradient fill | shape glow | rectangle shape | red to yellow gradient | glow size 5 points | Excel shape formatting
// Common Searches: Aspose.Cells add gradient fill to shape | C# set glow effect on Excel shape | how to create rectangle with gradient in Aspose.Cells | gradient fill and glow size in Aspose.Cells .NET
// Developer Intent: Add a rectangle shape with a horizontal red‑to‑yellow gradient and a 5‑point yellow glow in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Highlight important sections in a financial dashboard with a colored, glowing rectangle. | Create visually distinct headers for report pages in generated spreadsheets. | Design call‑out boxes that draw attention to key metrics across multiple worksheets.
// AI Prompts: Generate code to change the gradient direction to vertical and increase the glow size to 10 points. | Provide an example of a three‑color gradient with a semi‑transparent glow applied to a shape. | Explain how to reuse the same gradient and glow configuration for several shapes in a workbook.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to add a rectangle shape to the first worksheet, apply a horizontal red‑to‑yellow two‑color gradient fill, set a 5‑point yellow glow, and save the workbook with Aspose.Cells for .NET.
class GradientGlowDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);

        // Apply a two‑color gradient fill (red → yellow) to the shape
        shape.Fill.FillType = FillType.Gradient;
        GradientFill gradientFill = shape.Fill.GradientFill;
        gradientFill.SetTwoColorGradient(
            Color.Red,          // First color
            Color.Yellow,       // Second color
            GradientStyleType.Horizontal, // Gradient direction
            1);                 // Variant

        // Configure the glow effect: size 5 points, color yellow
        shape.Glow.Size = 5;
        shape.Glow.Color = workbook.CreateCellsColor();
        shape.Glow.Color.Color = Color.Yellow;

        // Save the workbook
        workbook.Save("GradientGlowDemo.xlsx");
    }
}
