// Title: C# Example: Change Accent3 Theme Color, Apply a Preset Theme Gradient to a Shape, and Verify Fill with Aspose.Cells
// Description: Demonstrates how to load an Excel workbook using Aspose.Cells for .NET, set the Accent3 theme color to orange, add a rectangle shape, apply a medium preset theme gradient that references Accent3, confirm the shape uses a gradient fill, add text, and save the modified file.
// Keywords: Aspose.Cells | C# theme color | Accent3 | preset theme gradient | gradient fill shape | Excel workbook theme | Aspose.Cells .NET example | shape fill verification | SetThemeColor | GetThemeColor
// Common Searches: how to change Accent3 theme color Aspose.Cells C# | apply preset theme gradient to shape Aspose.Cells | verify gradient fill type of a shape in Aspose.Cells | load and save workbook after theme modification Aspose.Cells | C# Aspose.Cells example for theme gradients
// Developer Intent: Update the workbook’s Accent3 theme color, apply a preset gradient that follows the theme, and confirm the shape’s fill type before saving.
// Use Cases: Enforce corporate branding by programmatically adjusting theme colors across multiple workbooks. | Create theme‑aware shapes that automatically adopt consistent gradient styles for reports and dashboards. | Automated quality checks to ensure visual elements reflect theme changes before distribution.
// AI Prompts: Generate C# code with Aspose.Cells to set the Accent2 theme color to a custom RGB value and apply a two‑color linear gradient to an ellipse shape. | Write a unit test that loads a workbook, changes Accent3, adds a rectangle with a preset theme gradient, and asserts that FillType equals Gradient. | Provide a step‑by‑step guide for updating several theme colors and assigning matching preset gradients to different shapes in an Aspose.Cells workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to load an Excel workbook using Aspose.Cells for .NET, set the Accent3 theme color to orange, add a rectangle shape, apply a medium preset theme gradient that references Accent3, confirm the shape uses a gradient fill, add text, and save the modified file.
class ThemeGradientDemo
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Change the theme's Accent3 color to a solid base color (e.g., Orange)
        workbook.SetThemeColor(ThemeColorType.Accent3, Color.Orange);

        // Verify that the theme color was updated
        Color accent3Color = workbook.GetThemeColor(ThemeColorType.Accent3);
        Console.WriteLine($"Accent3 theme color set to: {accent3Color}");

        // Add a rectangle shape to demonstrate a gradient that uses the Accent3 theme color
        Shape rect = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 200, 100);
        // Set the fill type to gradient so we can access GradientFill
        rect.Fill.FillType = FillType.Gradient;

        // Apply a preset theme gradient that references Accent3
        rect.Fill.GradientFill.SetPresetThemeGradient(
            PresetThemeGradientType.MediumGradient,   // gradient type
            ThemeColorType.Accent3);                  // theme color to base the gradient on

        // Verify that the shape's fill is a gradient
        bool isShapeGradient = rect.Fill.FillType == FillType.Gradient;
        Console.WriteLine($"Rectangle shape uses gradient fill: {isShapeGradient}");

        // Optionally, write some text inside the shape to see the effect
        rect.Text = "Accent3 Gradient";

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
