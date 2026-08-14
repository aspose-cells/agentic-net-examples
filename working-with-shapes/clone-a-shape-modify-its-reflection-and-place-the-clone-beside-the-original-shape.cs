// Title: Clone a Shape, Adjust Its Reflection, and Position It Beside the Original with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a rectangle shape, apply a half‑reflection, clone the shape using AddCopy, shift the copy five columns to the right, and set a different reflection effect on the cloned shape before saving the file.
// Keywords: Aspose.Cells shape cloning | C# AddCopy shape | reflection effect Aspose.Cells | duplicate rectangle shape .NET | modify shape reflection C# | Aspose.Cells worksheet graphics | Excel shape copy code
// Common Searches: Aspose.Cells clone shape and change reflection | AddCopy method example C# | how to set reflection effect on a copied shape in Aspose.Cells | move cloned shape to another column Aspose.Cells | C# code for shape duplication with different reflection
// Developer Intent: The developer needs to duplicate an existing shape, apply a new reflection style to the copy, and place the duplicate next to the original in an Excel worksheet.
// Use Cases: Create side‑by‑side design samples with distinct reflection styles for a presentation. | Generate a template where a logo is duplicated with varied reflections to showcase branding options. | Automate visual depth in reports by copying shapes and assigning custom reflection parameters.
// AI Prompts: Write C# code that uses Aspose.Cells to clone a rectangle shape, move the clone five columns right, and apply a full reflection effect. | Explain the parameters of the AddCopy method in Aspose.Cells and how they control the cloned shape's position. | Show how to access and modify advanced reflection properties such as Direction and FadeDirection after cloning a shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeCloneReflection
{
    // Demonstrates how to create a workbook, add a rectangle shape, apply a half‑reflection, clone the shape using AddCopy, shift the copy five columns to the right, and set a different reflection effect on the cloned shape before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add an original rectangle shape
            // Parameters: upper left row, upper left row offset, upper left column, upper left column offset, width, height
            Shape originalShape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 120, 80);

            // Optionally set a reflection on the original shape for visual reference
            ReflectionEffect originalReflection = originalShape.Reflection;
            originalReflection.Type = ReflectionEffectType.HalfReflectionTouching;
            originalReflection.Transparency = 0.3;
            originalReflection.Size = 60;
            originalReflection.Blur = 2;
            originalReflection.Distance = 5;

            // Clone the original shape and place the copy beside it
            // Place the copy 5 columns to the right of the original shape
            int newTopRow = 2;          // same top row
            int newTop = 0;             // same vertical offset
            int newLeftColumn = 7;      // shift right (original left column was 2)
            int newLeft = 0;            // same horizontal offset
            Shape clonedShape = worksheet.Shapes.AddCopy(originalShape, newTopRow, newTop, newLeftColumn, newLeft);

            // Modify the reflection of the cloned shape
            ReflectionEffect clonedReflection = clonedShape.Reflection;
            clonedReflection.Type = ReflectionEffectType.FullReflection4PtOffset;
            clonedReflection.Transparency = 0.6;
            clonedReflection.Size = 80;
            clonedReflection.Blur = 4;
            clonedReflection.Distance = 10;
            clonedReflection.Direction = 90;          // optional additional property
            clonedReflection.FadeDirection = 45;      // optional additional property

            // Save the workbook
            workbook.Save("ClonedShapeWithReflection.xlsx");
        }
    }
}
