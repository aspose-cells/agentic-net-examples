// Title: C# – Set Shape Reflection Size (30%) and Blur Radius (5 pt) with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a rectangle shape, and configure its ReflectionEffect by setting the Size to 30 % and the Blur radius to 5 points, then save the file as an .xlsx document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells reflection effect | C# shape blur radius | Aspose.Cells set reflection size | Aspose.Cells rectangle shape | ReflectionEffect Size property | ReflectionEffect Blur property | Excel shape styling Aspose | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set shape reflection size C# | How to set blur radius on shape reflection Aspose.Cells | ReflectionEffect Size and Blur in Aspose.Cells for .NET | Apply reflection to rectangle shape using Aspose.Cells | C# Aspose.Cells shape formatting tutorial
// Developer Intent: Apply a 30 % reflection size and a 5‑point blur radius to a rectangle shape in an Excel workbook programmatically with Aspose.Cells for .NET.
// Use Cases: Design visually appealing dashboard tiles with subtle reflective effects. | Highlight key metrics in financial reports by adding reflective shapes. | Standardize annotation styling across multiple worksheets by applying uniform reflection settings.
// AI Prompts: Generate C# code that sets a shape's ReflectionEffect Size to 30 % and Blur to 5 pt using Aspose.Cells. | Explain how to modify additional reflection properties (e.g., transparency, distance) after setting size and blur. | Show how to loop through all shapes on a worksheet and apply the same reflection size and blur settings.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a rectangle shape, and configure its ReflectionEffect by setting the Size to 30 % and the Blur radius to 5 points, then save the file as an .xlsx document using Aspose.Cells for .NET.
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

        // Set reflection size to 30 points (percentage) and blur radius to 5 points
        reflection.Size = 30;   // end position along the alpha gradient ramp (percentage)
        reflection.Blur = 5;    // blur radius in points

        // Save the workbook with the applied reflection effect
        workbook.Save("ReflectionShapeDemo.xlsx");
    }
}
