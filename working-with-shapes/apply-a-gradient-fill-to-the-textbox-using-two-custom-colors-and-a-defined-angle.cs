// Title: Aspose.Cells .NET – Apply a Two‑Color Linear Gradient with Custom Angle to a TextBox Shape
// Description: Create a workbook, add a TextBox shape, set its FillType to Gradient, define two ARGB colors, configure a 45° linear gradient, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells gradient fill | textbox shape gradient .NET | two color linear gradient Aspose | custom gradient angle Excel | shape fill color programming | Aspose.Cells C# gradient example
// Common Searches: Aspose.Cells set linear gradient on textbox | C# apply custom angle gradient to shape | two‑color gradient fill Aspose.Cells example | how to change gradient angle in Excel shape .NET | gradient fill textbox Aspose.Cells tutorial
// Developer Intent: Add a TextBox shape to a worksheet and style it with a two‑color linear gradient at a specific angle.
// Use Cases: Highlight key metrics in a report with an orange‑to‑blue gradient textbox. | Design an interactive dashboard where each section uses distinct gradient angles for visual separation. | Programmatically style form fields in an Excel template by applying custom gradient fills to TextBox controls.
// AI Prompts: Generate C# code to apply a three‑color gradient with a custom angle to a shape in Aspose.Cells. | Show how to change an existing textbox gradient from horizontal to diagonal using Aspose.Cells for .NET. | Explain how to read and modify the Angle property of a GradientFill after it has been applied.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Create a workbook, add a TextBox shape, set its FillType to Gradient, define two ARGB colors, configure a 45° linear gradient, and save the file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
        Shape textBox = sheet.Shapes.AddTextBox(1, 0, 10, 10, 200, 100);
        textBox.Text = "Gradient TextBox";

        // Set the fill type of the textbox to Gradient so we can access GradientFill properties
        textBox.Fill.FillType = FillType.Gradient;

        // Retrieve the GradientFill object associated with the textbox
        GradientFill gradientFill = textBox.Fill.GradientFill;

        // Define two custom colors for the gradient
        Color customColor1 = Color.FromArgb(255, 255, 200, 0);   // Orange
        Color customColor2 = Color.FromArgb(255, 0, 120, 215);   // Blue

        // Configure the gradient as a linear fill with a specific angle (e.g., 45 degrees)
        gradientFill.SetGradient(GradientFillType.Linear, 45, GradientDirectionType.FromUpperLeftCorner);
        // Apply the two‑color gradient using the custom colors
        gradientFill.SetTwoColorGradient(customColor1, customColor2, GradientStyleType.Horizontal, 1);
        // Ensure the Angle property reflects the desired angle
        gradientFill.Angle = 45f;

        // Save the workbook to a file
        workbook.Save("GradientTextBox.xlsx");
    }
}
