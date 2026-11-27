// Title: Assign a PNG image stream to a shape's TextureFill using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set texture fill of a shape using PNG byte array | load PNG into memory stream and apply as shape fill in Aspose.Cells workbook | assign image data to Shape.Fill.TextureFill.ImageData in Aspose.Cells .NET | create rectangle shape with texture fill from PNG file using Aspose.Cells | Aspose.Cells texture fill tiling and scaling example C#
// Tags: Aspose.Cells texture fill from PNG byte array | C# shape texture fill Aspose.Cells | assign image data to Shape.Fill.TextureFill | rectangle shape texture fill Aspose.Cells | texture fill tiling scaling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureFillDemo
{
    // The example loads a PNG file into a byte array, creates a workbook, adds a rectangle shape, sets its Fill.FillType to Texture, assigns the PNG bytes to the shape's TextureFill.ImageData, optionally configures tiling and scaling, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the PNG image that will be used as texture.
            // Adjust this path as needed; the file must exist.
            string pngPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "texture.png");

            if (!File.Exists(pngPath))
            {
                Console.WriteLine($"Error: PNG image not found at '{pngPath}'.");
                return;
            }

            try
            {
                // Load the PNG image into a memory stream.
                byte[] pngBytes;
                using (FileStream fs = new FileStream(pngPath, FileMode.Open, FileAccess.Read))
                using (MemoryStream ms = new MemoryStream())
                {
                    fs.CopyTo(ms);
                    pngBytes = ms.ToArray();
                }

                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet.
                // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
                Shape shape = sheet.Shapes.AddRectangle(2, 0, 0, 100, 200, 100);

                // Set the fill type to Texture so we can assign a texture fill.
                shape.Fill.FillType = FillType.Texture;

                // Get the TextureFill object.
                TextureFill textureFill = shape.Fill.TextureFill;

                // Assign the PNG image data to the texture fill.
                // The property that holds the image bytes is ImageData.
                textureFill.ImageData = pngBytes;

                // Optional: configure tiling and scaling.
                textureFill.IsTiling = true;          // Tile the picture across the shape.
                textureFill.Scale = 0.8;              // Scale the picture to 80% of its original size.
                textureFill.Transparency = 0.0;       // Fully opaque.

                // Save the workbook.
                string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TextureFillShape.xlsx");
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The rectangle shape now uses the PNG image as a texture fill.");
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
