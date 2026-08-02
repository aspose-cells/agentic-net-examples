// Title: Read a shape's shadow color and log its RGB values using Aspose.Cells for .NET
// Description: Shows how to access a worksheet shape's ShadowEffect.Color, convert it to System.Drawing.Color, and write the red, green, and blue components to the console (with optional workbook save).
// Keywords: Aspose.Cells shadow color | shape shadow RGB | read shape shadow Aspose | Excel shape formatting .NET | retrieve shadow effect color | Aspose.Cells shape styling
// Common Searches: Aspose.Cells get shape shadow color | How to obtain RGB of shape shadow in C# | Read shadow effect color of Excel shape | Aspose.Cells shape formatting shadow example
// Developer Intent: Extract the current shadow color of a worksheet shape and output its RGB components.
// Use Cases: Verify that programmatically created shapes use the intended shadow hue during automated report generation. | Debug visual inconsistencies in Excel files by logging shadow color values for each shape. | Generate an audit trail of shape formatting settings, including shadow color, for compliance documentation.
// AI Prompts: Write C# code with Aspose.Cells that reads a shape's shadow color and logs the RGB values to a file. | Provide a loop that iterates over all shapes in a worksheet and prints each shape's shadow color components. | Explain how to change a shape's shadow color based on its existing RGB values using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// Shows how to access a worksheet shape's ShadowEffect.Color, convert it to System.Drawing.Color, and write the red, green, and blue components to the console (with optional workbook save).
class ReadShadowColorDemo
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 100);

        // Set a shadow color for the shape (for demonstration purposes)
        ShadowEffect shadowEffect = shape.ShadowEffect;
        CellsColor shadowColor = workbook.CreateCellsColor();
        shadowColor.Color = Color.Green; // Example shadow color
        shadowEffect.Color = shadowColor;

        // Retrieve the current shadow color of the shape
        CellsColor currentShadowColor = shape.ShadowEffect.Color;
        Color sysColor = currentShadowColor.Color;

        // Log the RGB components of the shadow color
        Console.WriteLine($"Shadow Color - R: {sysColor.R}, G: {sysColor.G}, B: {sysColor.B}");

        // Save the workbook (optional)
        workbook.Save("ReadShadowColorDemo.xlsx");
    }
}
