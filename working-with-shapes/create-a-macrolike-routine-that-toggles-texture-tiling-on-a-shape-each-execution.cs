using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsDemo
{
    public class TextureTilingToggle
    {
        public static void Run()
        {
            try
            {
                // Path of the workbook that will store the shape
                string filePath = "ToggleTextureTiling.xlsx";

                // Load existing workbook if it exists, otherwise create a new one
                Workbook workbook = File.Exists(filePath) ? new Workbook(filePath) : new Workbook();

                // Work with the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Try to find a shape named "MyShape"
                Shape shape = null;
                foreach (Shape s in worksheet.Shapes)
                {
                    if (s.Name == "MyShape")
                    {
                        shape = s;
                        break;
                    }
                }

                // If the shape does not exist, create it and set an initial texture
                if (shape == null)
                {
                    // Add a rectangle shape
                    shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);
                    shape.Name = "MyShape";

                    // Configure texture fill
                    shape.Fill.FillType = FillType.Texture;
                    shape.Fill.TextureFill.Type = TextureType.BlueTissuePaper;

                    // Start with tiling disabled
                    shape.Fill.TextureFill.IsTiling = false;
                }

                // Toggle the IsTiling property
                shape.Fill.TextureFill.IsTiling = !shape.Fill.TextureFill.IsTiling;

                Console.WriteLine($"Texture tiling is now set to: {shape.Fill.TextureFill.IsTiling}");

                // Save the workbook (creates the file on first run, updates on subsequent runs)
                workbook.Save(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TextureTilingToggle.Run();
        }
    }
}