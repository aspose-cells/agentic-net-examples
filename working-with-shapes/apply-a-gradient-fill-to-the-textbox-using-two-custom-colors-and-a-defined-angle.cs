// Title: Add a Two‑Color Linear Gradient with Custom Angle to a TextBox Shape in Aspose.Cells for .NET
// Description: Shows how to create a workbook, insert a TextBox shape, set its FillType to Gradient, define a linear gradient at a custom angle, apply two custom colors, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | gradient fill | TextBox shape | linear gradient | custom angle | two‑color gradient | FillType | GradientFill | Excel shape styling
// Common Searches: Aspose.Cells set gradient fill for textbox | C# linear gradient with angle Aspose.Cells | two color gradient shape Aspose.Cells .NET | how to apply custom angle gradient to Excel shape using Aspose | gradient fill textbox Aspose.Cells example
// Developer Intent: Apply a linear gradient with a specific angle and two custom colors to a TextBox shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Design a report header with an angled orange‑to‑blue gradient textbox for visual impact. | Highlight key sections of a worksheet by giving shapes a custom‑angled two‑color gradient background. | Create dashboard widgets where the gradient direction matches surrounding chart elements.
// AI Prompts: Generate C# code that adds a rounded rectangle shape with a radial gradient using three custom colors in Aspose.Cells. | Show how to change a shape's gradient angle dynamically based on a cell value in Aspose.Cells for .NET. | Provide an example that applies a vertical two‑color gradient to a collection of shapes in a loop using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, insert a TextBox shape, set its FillType to Gradient, define a linear gradient at a custom angle, apply two custom colors, and save the file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape textBox = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 200, 100);
        textBox.Text = "Gradient TextBox";

        // Set the fill type of the text box to Gradient
        textBox.Fill.FillType = FillType.Gradient;

        // Obtain the GradientFill object from the shape's Fill
        GradientFill gradientFill = textBox.Fill.GradientFill;

        // Define a linear gradient with a specific angle (e.g., 45 degrees)
        gradientFill.SetGradient(GradientFillType.Linear, 45, GradientDirectionType.FromCenter);
        // Alternatively, you can set the Angle property directly
        // gradientFill.Angle = 45;

        // Apply a two‑color gradient using custom colors
        Color customColor1 = Color.FromArgb(255, 255, 128, 0); // Orange
        Color customColor2 = Color.FromArgb(255, 0, 128, 255); // Blue
        gradientFill.SetTwoColorGradient(
            customColor1,          // First color
            customColor2,          // Second color
            GradientStyleType.Horizontal, // Gradient style
            1);                    // Variant (1‑4)

        // Save the workbook with the gradient‑filled text box
        workbook.Save("TextboxGradientFill.xlsx");
    }
}
