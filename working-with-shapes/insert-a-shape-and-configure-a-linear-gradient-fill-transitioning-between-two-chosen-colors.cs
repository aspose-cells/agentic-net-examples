// Title: Aspose.Cells C# – Insert Rectangle Shape with Horizontal Two‑Color Linear Gradient Fill
// Description: Creates a new workbook, adds a rectangle shape to the first worksheet, sets its fill type to Gradient, defines a horizontal linear gradient, applies an orange‑to‑purple two‑color gradient, and saves the file as LinearGradientShape.xlsx.
// Keywords: Aspose.Cells add shape gradient | C# linear gradient shape | two color gradient rectangle Aspose.Cells | horizontal gradient fill shape | Aspose.Cells gradient fill example
// Common Searches: how to apply a horizontal linear gradient to a shape in Aspose.Cells C# | Aspose.Cells example two‑color gradient rectangle | set gradient direction and style on a shape Aspose.Cells | save workbook with gradient‑filled shapes C#
// Developer Intent: Add a rectangle shape, apply a horizontal two‑color linear gradient, and save the workbook.
// Use Cases: Create a banner with a smooth color transition for report headers. | Design a status indicator that shifts from orange to purple to convey state changes. | Build a custom background rectangle with a horizontal gradient for dashboards or charts.
// AI Prompts: Generate C# code to change the gradient angle to 45° for the rectangle shape in Aspose.Cells. | Show how to add an ellipse shape with a radial two‑color gradient using Aspose.Cells for .NET. | Explain how to read and modify the start and end colors of an existing shape's gradient fill.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a rectangle shape to the first worksheet, sets its fill type to Gradient, defines a horizontal linear gradient, applies an orange‑to‑purple two‑color gradient, and saves the file as LinearGradientShape.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a rectangle shape (row, column, upperLeftRow, upperLeftColumn, width, height)
        Shape shape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 300, 150);

        // Set the fill type of the shape to Gradient so we can access GradientFill properties
        shape.Fill.FillType = FillType.Gradient;

        // Obtain the GradientFill object
        GradientFill gradientFill = shape.Fill.GradientFill;

        // Define a linear gradient (angle = 0 for horizontal)
        gradientFill.SetGradient(GradientFillType.Linear, 0, GradientDirectionType.FromCenter);

        // Apply a two‑color gradient (first color, second color, style, variant)
        gradientFill.SetTwoColorGradient(
            Color.Orange,          // start color
            Color.Purple,          // end color
            GradientStyleType.Horizontal, // gradient style
            1);                    // variant (1‑4)

        // Save the workbook with the shape that has the linear gradient fill
        workbook.Save("LinearGradientShape.xlsx");
    }
}
