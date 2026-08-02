// Title: Clone a Shape and Change Its Reflection with Aspose.Cells for .NET (C#)
// Description: Shows how to add a rectangle to an Excel worksheet, apply a full reflection, duplicate the shape using AddCopy, reposition the copy beside the original, assign a different half‑reflection effect, and save the workbook.
// Keywords: Aspose.Cells | .NET | C# | shape cloning | AddCopy method | reflection effect | Excel shape | duplicate shape | rectangle shape | programmatic Excel graphics
// Common Searches: Aspose.Cells clone shape C# | How to copy a shape to another column in Excel using Aspose.Cells | Set reflection effect on Excel shape with Aspose.Cells | AddCopy method example Aspose.Cells | Change shape reflection after cloning Aspose.Cells
// Developer Intent: The developer wants to programmatically duplicate an existing shape in an Excel sheet, move the copy to a new location, and apply a different reflection style to the duplicated shape.
// Use Cases: Design templates that compare two visual elements with distinct reflections. | Automated generation of marketing sheets where one shape shows a full mirror and another a partial mirror. | Creating decorative dashboards that require side‑by‑side shapes with customized reflection properties.
// AI Prompts: Generate C# code that copies a rectangle shape in Aspose.Cells and sets a different ReflectionEffect on the copy. | Show how to use the AddCopy method to place a cloned shape in another column and modify its reflection parameters. | Explain step‑by‑step how to clone a shape, reposition it, and change its reflection effect with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to add a rectangle to an Excel worksheet, apply a full reflection, duplicate the shape using AddCopy, reposition the copy beside the original, assign a different half‑reflection effect, and save the workbook.
class ShapeCloneReflectionDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add an original rectangle shape
        // Parameters: upper left row, upper left row offset, upper left column, upper left column offset, width, height
        Shape original = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 120, 80);

        // Set reflection properties for the original shape
        ReflectionEffect origReflection = original.Reflection;
        origReflection.Type = ReflectionEffectType.FullReflection4PtOffset;
        origReflection.Transparency = 0.3;
        origReflection.Size = 60;
        origReflection.Blur = 2;
        origReflection.Distance = 5;

        // Clone the original shape and place the copy beside it
        // Place the clone at the same top row (2) but a different left column (7)
        Shape clone = worksheet.Shapes.AddCopy(original, 2, 0, 7, 0);

        // Modify reflection properties of the cloned shape
        ReflectionEffect cloneReflection = clone.Reflection;
        cloneReflection.Type = ReflectionEffectType.HalfReflectionTouching;
        cloneReflection.Transparency = 0.6;
        cloneReflection.Size = 40;
        cloneReflection.Blur = 1;
        cloneReflection.Distance = 3;

        // Save the workbook with the shapes
        workbook.Save("ShapeCloneReflectionDemo.xlsx");
    }
}
