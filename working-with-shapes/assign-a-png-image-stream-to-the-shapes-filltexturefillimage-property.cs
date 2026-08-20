// Title: Set a PNG texture fill for a shape using MemoryStream in Aspose.Cells for .NET
// Description: Learn how to load a PNG file into a MemoryStream, assign it to a shape's Fill.TextureFill.ImageData, and optionally configure tiling or scaling with Aspose.Cells for .NET. The example creates a workbook, adds a rectangle, applies the texture, and saves the Excel file.
// Keywords: Aspose.Cells texture fill | C# shape fill PNG | Fill.TextureFill.ImageData | MemoryStream PNG Aspose | Excel shape texture fill .NET | Aspose.Cells ImageData stream | shape texture tiling Aspose | apply PNG to shape Aspose.Cells
// Common Searches: Aspose.Cells set PNG texture fill for shape | C# assign image stream to Shape.Fill.TextureFill.ImageData | load PNG into MemoryStream for Excel shape fill | how to enable tiling on texture fill Aspose.Cells | Aspose.Cells shape fill from byte array
// Developer Intent: Apply a PNG image as a texture fill to a worksheet shape by using a MemoryStream with Aspose.Cells for .NET.
// Use Cases: Create a rectangle (or any shape) and apply a custom PNG texture using a byte array from a MemoryStream. | Load a PNG file at runtime, convert it to bytes, and set it as the Fill.TextureFill.ImageData for dynamic Excel reports. | Adjust texture properties such as IsTiling or Scale after assigning the PNG to control repetition and size within the shape.
// AI Prompts: Generate C# code that reads a PNG file into a MemoryStream and sets Shape.Fill.TextureFill.ImageData in Aspose.Cells. | Show how to enable tiling and set the scale factor for a texture fill after assigning a PNG image stream. | Explain error handling for missing image files when applying a texture fill to a shape with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureFillDemo
{
    // Learn how to load a PNG file into a MemoryStream, assign it to a shape's Fill.TextureFill.ImageData, and optionally configure tiling or scaling with Aspose.Cells for .NET. The example creates a workbook, adds a rectangle, applies the texture, and saves the Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 0, 150, 100);

                // Set the fill type of the shape to Texture
                shape.Fill.FillType = FillType.Texture;

                // Get the TextureFill object associated with the shape
                TextureFill textureFill = shape.Fill.TextureFill;

                string imagePath = "image.png";

                if (File.Exists(imagePath))
                {
                    // Load a PNG image into a memory stream
                    using (FileStream fileStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                    using (MemoryStream pngStream = new MemoryStream())
                    {
                        fileStream.CopyTo(pngStream);
                        // Assign the PNG image data to the texture fill
                        textureFill.ImageData = pngStream.ToArray();
                    }
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Texture fill will be skipped.");
                }

                // Optionally enable tiling or adjust other texture properties
                // textureFill.IsTiling = true;
                // textureFill.Scale = 0.5;

                // Save the workbook to a file
                workbook.Save("TextureFillDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
