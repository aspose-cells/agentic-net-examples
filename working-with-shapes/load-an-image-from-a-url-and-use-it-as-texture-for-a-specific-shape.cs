// Title: Download an image from a URL and use it as a texture fill for a rectangle shape in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set rectangle shape fill to image downloaded from web | How to apply a web image as texture fill to a shape using Aspose.Cells .NET | Aspose.Cells texture fill shape with image bytes from HttpClient | Load external PNG as shape texture in Excel workbook with Aspose.Cells | Apply tiling and scaling to shape texture using Aspose.Cells C#
// Tags: texture fill rectangle shape Aspose.Cells | download image bytes HttpClient C# | assign image data to TextureFill Aspose.Cells | shape fill tiling scaling Aspose.Cells | save workbook with textured shape Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example downloads an image from a public URL using HttpClient, creates a new workbook, adds a rectangle shape, sets its fill type to texture, assigns the downloaded image bytes to the shape's TextureFill, optionally enables tiling and scaling, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        // URL of the image to be used as texture.
        // Replace with a valid, publicly accessible image URL.
        const string imageUrl = "https://example.com/sample-image.png";

        // Download the image data.
        byte[] imageData;
        using (HttpClient client = new HttpClient())
        {
            try
            {
                HttpResponseMessage response = client.GetAsync(imageUrl).Result;
                response.EnsureSuccessStatusCode();
                imageData = response.Content.ReadAsByteArrayAsync().Result;
                Console.WriteLine($"Image downloaded successfully ({imageData.Length} bytes).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error downloading image from '{imageUrl}': {ex.Message}");
                return;
            }
        }

        // Create a new workbook and get the first worksheet.
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape that will receive the texture.
        // Parameters: upper left row, top offset (pixels), upper left column, left offset (pixels), height (pixels), width (pixels)
        Shape rectangle = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 200, 300);

        // Set the fill type of the shape to texture.
        rectangle.Fill.FillType = FillType.Texture;

        // Obtain the TextureFill object and assign the downloaded image.
        TextureFill textureFill = rectangle.Fill.TextureFill;
        textureFill.ImageData = imageData;

        // Optional: configure tiling and scaling.
        textureFill.IsTiling = true;   // Tile the picture across the shape.
        textureFill.Scale = 0.8;       // Scale the texture to 80% of its original size.

        // Save the workbook.
        const string outputFile = "ShapeWithTexture.xlsx";
        try
        {
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
