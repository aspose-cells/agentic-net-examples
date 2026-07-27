// Title: Assign a PNG stream to Shape.Fill.TextureFill.ImageData with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to load a PNG file into a MemoryStream, assign its byte array to a shape's TextureFill.ImageData, enable tiling, adjust scaling, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells texture fill PNG | Shape.Fill.TextureFill.ImageData | C# memory stream shape fill | Aspose.Cells set shape texture from stream | tiling scaling texture fill .NET
// Common Searches: Aspose.Cells set PNG as texture fill from stream | C# assign image bytes to Shape.Fill.TextureFill | How to use MemoryStream for shape texture in Aspose.Cells | Enable tiling for texture fill Aspose.Cells | Scale texture fill image Aspose.Cells C#
// Developer Intent: Apply a PNG image supplied via a stream as the texture fill of a worksheet shape.
// Use Cases: Create a rectangle shape and fill it with a PNG pattern loaded from a file or uploaded by a user. | Apply a tiled PNG texture with custom scaling to multiple shapes in the same worksheet. | Generate workbooks where shape fills are driven by in‑memory images rather than external files.
// AI Prompts: Generate C# code that loads a PNG into a MemoryStream and sets Shape.Fill.TextureFill.ImageData in Aspose.Cells, including tiling and scaling options. | Show how to reuse a single PNG byte array for the texture fill of several shapes in an Aspose.Cells workbook. | Explain the steps to configure TextureFill.IsTiling and TextureFill.Scale after assigning ImageData from a stream.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureFillDemo
{
    // Demonstrates how to load a PNG file into a MemoryStream, assign its byte array to a shape's TextureFill.ImageData, enable tiling, adjust scaling, and save the workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 0, 200, 100);

            // Set the fill type of the shape to texture so that TextureFill can be used
            shape.Fill.FillType = FillType.Texture;

            // Obtain the TextureFill object from the shape's Fill
            TextureFill textureFill = shape.Fill.TextureFill;

            // Load a PNG image into a memory stream (replace the path with your actual PNG file)
            using (FileStream fileStream = new FileStream("sample.png", FileMode.Open, FileAccess.Read))
            using (MemoryStream pngStream = new MemoryStream())
            {
                fileStream.CopyTo(pngStream);
                // Assign the PNG image data to the texture fill
                textureFill.ImageData = pngStream.ToArray();
            }

            // Optionally enable tiling or adjust other texture properties
            textureFill.IsTiling = true;
            textureFill.Scale = 0.8; // 80% scaling

            // Save the workbook to a file
            workbook.Save("TextureFillWithPng.xlsx");
        }
    }
}
