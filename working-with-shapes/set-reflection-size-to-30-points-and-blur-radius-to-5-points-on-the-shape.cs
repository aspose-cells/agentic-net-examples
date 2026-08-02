// Title: Aspose.Cells for .NET: Set Shape Reflection Size (30 pt) and Blur Radius (5 pt) in C#
// Description: C# example that creates a workbook, adds a rectangle shape, accesses its ReflectionEffect, sets the reflection size to 30 points and blur radius to 5 points, and saves the file as ShapeReflectionSettings.xlsx.
// Keywords: Aspose.Cells shape reflection | C# reflection size 30 | blur radius 5 Aspose.Cells | Excel shape effects .NET | set shape reflection properties | Aspose.Cells rectangle reflection | reflection effect API | Excel workbook styling C# | Aspose.Cells graphics options | programmatic shape formatting
// Common Searches: how to set reflection size of a shape in Aspose.Cells C# | Aspose.Cells blur radius for shape reflection | apply reflection effect to Excel shape using .NET | change shape reflection properties programmatically | Aspose.Cells shape styling examples
// Developer Intent: Programmatically apply a reflection effect with a 30‑point size and 5‑point blur to a shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Design polished report headers with a subtle reflective rectangle for visual emphasis. | Generate product catalogs where every image placeholder shares identical reflection settings for brand consistency. | Automate template creation that requires uniform reflection effects across multiple shapes in different worksheets.
// AI Prompts: Show how to modify additional reflection parameters such as offset and transparency for a shape with Aspose.Cells in C#. | Provide code to apply the same 30‑point reflection size and 5‑point blur to all shapes on a worksheet using Aspose.Cells. | Explain how to read, update, or remove the reflection effect of an existing shape in an Excel file with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, adds a rectangle shape, accesses its ReflectionEffect, sets the reflection size to 30 points and blur radius to 5 points, and saves the file as ShapeReflectionSettings.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, width, height
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);

        // Access the reflection effect of the shape
        ReflectionEffect reflection = shape.Reflection;

        // Set reflection size to 30 points (percentage of the gradient ramp)
        reflection.Size = 30;

        // Set reflection blur radius to 5 points
        reflection.Blur = 5;

        // Save the workbook with the applied reflection settings
        workbook.Save("ShapeReflectionSettings.xlsx");
    }
}
