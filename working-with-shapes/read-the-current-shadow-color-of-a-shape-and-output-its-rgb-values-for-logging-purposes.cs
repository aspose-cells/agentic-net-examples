// Title: Read a shape's shadow color and log RGB values with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a rectangle shape, assigns a shadow color, reads the ShadowEffect.Color, extracts the System.Drawing.Color, and writes the R, G, B components to the console. It then saves the file, reloads it, and confirms that the shadow color persists.
// Keywords: Aspose.Cells | C# | shape shadow color | ShadowEffect | RGB values | read shadow color | retrieve shape styling | Excel shape formatting | Aspose.Cells API | worksheet shape
// Common Searches: Aspose.Cells get shape shadow color C# | How to read RGB of a shape's shadow in Aspose.Cells | ShadowEffect.Color example Aspose.Cells | Verify shape shadow after saving workbook Aspose.Cells | C# read shape formatting Aspose.Cells
// Developer Intent: Retrieve the current shadow color of a worksheet shape and output its RGB components.
// Use Cases: Debug visual formatting by logging shadow colors of generated shapes. | Validate that custom shadow styling survives workbook save and reload operations. | Create a style audit report that records RGB values of shape shadows across multiple sheets.
// AI Prompts: Generate C# code that reads a shape's ShadowEffect.Color and prints its RGB values using Aspose.Cells. | Show how to confirm that a shape's shadow color remains unchanged after saving and reopening an Excel file with Aspose.Cells. | Explain the steps to set, retrieve, and log the RGB components of a shape's shadow in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds a rectangle shape, assigns a shadow color, reads the ShadowEffect.Color, extracts the System.Drawing.Color, and writes the R, G, B components to the console. It then saves the file, reloads it, and confirms that the shadow color persists.
class ReadShadowColorDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 100);

        // Configure shadow color (for demonstration)
        ShadowEffect shadowEffect = shape.ShadowEffect;
        CellsColor shadowColor = workbook.CreateCellsColor();
        shadowColor.Color = Color.Blue;               // Set desired shadow color
        shadowEffect.Color = shadowColor;              // Apply to the shape's shadow

        // Read the current shadow color and output its RGB components
        Color currentColor = shape.ShadowEffect.Color.Color;
        Console.WriteLine($"Shadow Color RGB: {currentColor.R}, {currentColor.G}, {currentColor.B}");

        // Save the workbook
        string filePath = "ReadShadowColorDemo.xlsx";
        workbook.Save(filePath);

        // Load the workbook back to verify the saved shadow color
        Workbook loadedWorkbook = new Workbook(filePath);
        Shape loadedShape = loadedWorkbook.Worksheets[0].Shapes[0];
        Color loadedColor = loadedShape.ShadowEffect.Color.Color;
        Console.WriteLine($"Loaded Shadow Color RGB: {loadedColor.R}, {loadedColor.G}, {loadedColor.B}");
    }
}
