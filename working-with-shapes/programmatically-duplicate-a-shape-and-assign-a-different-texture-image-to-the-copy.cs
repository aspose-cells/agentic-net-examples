// Title: How to duplicate an AutoShape and assign a different texture fill using Aspose.Cells for .NET (C#)
// Common Searches: aspnet cells copy auto shape and change texture image c# | c# use Shape.AddCopy to clone shape with offset in Aspose.Cells | apply a different PNG texture to a duplicated shape Aspose.Cells example | set texture fill scaling and tiling on copied AutoShape in Aspose.Cells | duplicate rectangle shape and assign new fill image using Aspose.Cells .NET
// Tags: Shape.AddCopy method Aspose.Cells | texture fill for AutoShape C# | duplicate shape with different image Aspose.Cells | apply PNG texture to shape Aspose.Cells | shape fill scaling and tiling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeDuplicateWithTexture
{
    // The example creates a workbook, adds a rectangle AutoShape, applies a PNG texture as the fill, duplicates the shape using the AddCopy overload with a positional offset, then assigns a different PNG texture and distinct fill settings (tiling, scaling) to the copy before saving the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to texture images – adjust these paths to point to real image files on your machine.
            string texturePath1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "texture1.png");
            string texturePath2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "texture2.png");

            // Verify that the texture files exist before proceeding.
            if (!File.Exists(texturePath1))
            {
                Console.WriteLine($"Error: Texture image not found at \"{texturePath1}\"");
                return;
            }

            if (!File.Exists(texturePath2))
            {
                Console.WriteLine($"Error: Texture image not found at \"{texturePath2}\"");
                return;
            }

            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and add a rectangle shape.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape (AutoShapeType.Rectangle) at row 2, column 2.
                // Parameters: shape type, upper left row, top offset, upper left column, left offset, height, width
                Shape originalShape = sheet.Shapes.AddAutoShape(
                    AutoShapeType.Rectangle,
                    2, 0,   // top row, top offset (pixels)
                    2, 0,   // left column, left offset (pixels)
                    100,    // height (pixels)
                    200);   // width (pixels)

                // Apply the first texture to the original shape.
                originalShape.Fill.FillType = FillType.Texture;
                originalShape.Fill.TextureFill.ImageData = File.ReadAllBytes(texturePath1);
                originalShape.Fill.TextureFill.IsTiling = true; // optional tiling
                originalShape.Name = "OriginalShape";

                Console.WriteLine("Original shape created and texture applied.");

                // -----------------------------------------------------------------
                // 2. Duplicate the shape.
                // -----------------------------------------------------------------
                // Use AddCopy overload that takes the source shape and new position.
                Shape copiedShape = sheet.Shapes.AddCopy(
                    originalShape,
                    originalShape.UpperLeftRow + 5, // new top row
                    0,                               // top offset (pixels)
                    originalShape.UpperLeftColumn + 5, // new left column
                    0);                               // left offset (pixels)

                copiedShape.Name = "CopiedShape";

                // -----------------------------------------------------------------
                // 3. Assign a different texture image to the copied shape.
                // -----------------------------------------------------------------
                copiedShape.Fill.FillType = FillType.Texture;
                copiedShape.Fill.TextureFill.ImageData = File.ReadAllBytes(texturePath2);
                copiedShape.Fill.TextureFill.IsTiling = false; // demonstrate a different setting
                copiedShape.Fill.TextureFill.Scale = 0.7; // optional scaling

                Console.WriteLine("Copied shape created and new texture applied.");

                // -----------------------------------------------------------------
                // 4. Save the workbook.
                // -----------------------------------------------------------------
                string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ShapeDuplicateWithTexture.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
