// Title: C# example: Remove all shapes from a worksheet and add a rectangle with tiled texture fill and custom text using Aspose.Cells
// Common Searches: how to clear all shapes from an Excel worksheet using Aspose.Cells in C# | add a rectangle shape with tiled image fill in Aspose.Cells .NET | set custom text inside a shape created with Aspose.Cells C# | apply texture fill with tiling options to a shape in Aspose.Cells | save workbook after modifying shapes with Aspose.Cells C# example
// Tags: remove all shapes worksheet Aspose.Cells | add rectangle shape texture fill Aspose.Cells | tiled texture fill shape Aspose.Cells | assign custom text to shape Aspose.Cells | save workbook after shape changes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeDemo
{
    // The sample loads a workbook, validates the source and texture image files, removes every existing shape from the first worksheet, creates a new rectangle shape at a specific location, sets custom text inside the shape, applies a tiled texture fill using the supplied image with scaling options, and saves the updated workbook to the specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            // Expected arguments:
            // args[0] - path to the source workbook (e.g., "input.xlsx")
            // args[1] - path to the texture image (e.g., "texture.png")
            // args[2] - path for the output workbook (e.g., "output.xlsx")
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: AsposeCellsShapeDemo <sourceWorkbook> <textureImage> <outputWorkbook>");
                return;
            }

            string sourceWorkbookPath = args[0];
            string textureImagePath = args[1];
            string outputWorkbookPath = args[2];

            // Validate source workbook
            if (!File.Exists(sourceWorkbookPath))
            {
                Console.WriteLine($"Error: Source workbook not found at \"{sourceWorkbookPath}\"");
                return;
            }

            // Validate texture image
            if (!File.Exists(textureImagePath))
            {
                Console.WriteLine($"Error: Texture image not found at \"{textureImagePath}\"");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(sourceWorkbookPath);
                // Work with the first worksheet (you can change the index as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Remove all existing shapes from the worksheet
                // -------------------------------------------------
                ShapeCollection shapes = worksheet.Shapes;
                // Remove from the end to avoid index shifting
                for (int i = shapes.Count - 1; i >= 0; i--)
                {
                    shapes.RemoveAt(i);
                }
                Console.WriteLine($"All existing shapes removed. Remaining shape count: {shapes.Count}");

                // -------------------------------------------------
                // 2. Add a new rectangle shape
                // -------------------------------------------------
                // Parameters: upper left row, upper left column, upper left offsetX, offsetY, width, height (in pixels)
                int upperLeftRow = 2;
                int upperLeftColumn = 1;
                int upperLeftOffsetX = 0;
                int upperLeftOffsetY = 0;
                int shapeWidth = 300;   // pixels
                int shapeHeight = 150;  // pixels

                Shape newShape = worksheet.Shapes.AddRectangle(
                    upperLeftRow,
                    upperLeftColumn,
                    upperLeftOffsetX,
                    upperLeftOffsetY,
                    shapeWidth,
                    shapeHeight);

                // Set custom text inside the shape
                newShape.Text = "Custom Text – Aspose.Cells";

                // -------------------------------------------------
                // 3. Apply tiled texture fill to the shape
                // -------------------------------------------------
                // Load texture image bytes
                byte[] textureBytes = File.ReadAllBytes(textureImagePath);

                // Configure fill as texture
                newShape.Fill.FillType = FillType.Texture;
                TextureFill textureFill = newShape.Fill.TextureFill;

                // Assign image data
                textureFill.ImageData = textureBytes;

                // Enable tiling
                textureFill.IsTiling = true;

                // Optional: adjust tiling options (scale, offset, etc.)
                TilePicOption tileOption = new TilePicOption
                {
                    ScaleX = 0.5,   // 50% horizontal scale
                    ScaleY = 0.5,   // 50% vertical scale
                    OffsetX = 0,
                    OffsetY = 0
                };
                textureFill.TilePicOption = tileOption;

                Console.WriteLine("New shape added with tiled texture and custom text.");

                // -------------------------------------------------
                // 4. Save the modified workbook
                // -------------------------------------------------
                workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputWorkbookPath}\"");
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
