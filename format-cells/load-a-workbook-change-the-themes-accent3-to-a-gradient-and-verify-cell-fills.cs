// Title: Set Accent3 Theme Gradient on a Shape and Verify Fill Type with Aspose.Cells for .NET
// Description: Loads an Excel workbook, reads the current Accent3 theme color, adds a rectangle shape, applies a preset Accent3 gradient fill, confirms the shape's FillType is Gradient, and saves the updated file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | theme gradient | Accent3 | preset theme gradient | shape fill gradient | GradientFill | SetPresetThemeGradient | Excel workbook theme color | AddRectangle shape
// Common Searches: Aspose.Cells set Accent3 gradient | apply preset theme gradient Aspose.Cells .NET | check shape fill type Aspose.Cells | retrieve theme colors Aspose.Cells | add rectangle shape with gradient fill Aspose.Cells
// Developer Intent: Programmatically apply a preset Accent3 gradient to a shape, verify the fill type, and persist the changes in an Excel workbook.
// Use Cases: Demonstrate the Accent3 theme color by applying a preset gradient to a rectangle shape. | Log the original Accent3 color before modification for audit or comparison. | Validate that a shape’s FillType is Gradient before finalizing the workbook.
// AI Prompts: Generate C# code with Aspose.Cells that adds a shape and sets its fill to a custom gradient based on the Accent3 theme color. | Create a method to change the workbook’s Accent3 theme color to a specific RGB value and automatically update all existing gradient fills. | Explain how to iterate through all shapes in a worksheet, verify each shape’s FillType, and log the results using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, reads the current Accent3 theme color, adds a rectangle shape, applies a preset Accent3 gradient fill, confirms the shape's FillType is Gradient, and saves the updated file using Aspose.Cells for .NET.
class ThemeAccent3GradientDemo
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve and display the current Accent3 theme color
        Color originalAccent3 = workbook.GetThemeColor(ThemeColorType.Accent3);
        Console.WriteLine($"Original Accent3 color: {originalAccent3}");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape that will demonstrate the gradient fill
        // Parameters: upper left row, upper left column, upper left offsetX, upper left offsetY, width, height
        Shape rect = sheet.Shapes.AddRectangle(2, 1, 0, 0, 200, 100);

        // Set the shape's fill type to Gradient so we can work with GradientFill
        rect.Fill.FillType = FillType.Gradient;

        // Obtain the GradientFill object
        GradientFill gradientFill = rect.Fill.GradientFill;

        // Apply a preset theme gradient that uses the Accent3 theme color
        // Here we choose a MediumGradient; you can select any PresetThemeGradientType value
        gradientFill.SetPresetThemeGradient(
            PresetThemeGradientType.MediumGradient,
            ThemeColorType.Accent3);

        // Verify that the shape's fill is indeed a gradient
        bool isGradient = rect.Fill.FillType == FillType.Gradient;
        Console.WriteLine($"Shape fill is gradient: {isGradient}");

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
