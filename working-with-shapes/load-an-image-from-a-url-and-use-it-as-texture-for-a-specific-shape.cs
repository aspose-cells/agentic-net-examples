// Title: Apply a Web‑Downloaded Image as a Tiled Texture Fill to a Shape in Aspose.Cells (C#)
// Description: Creates a workbook, adds a rectangle shape, downloads an image (with local fallback), sets the shape's FillType to Texture, enables tiling, scales the texture to 50 % and saves the file.
// Keywords: Aspose.Cells texture fill | C# shape fill from URL | download image for Excel shape | Aspose.Cells FillType.Texture | image fallback Aspose.Cells | tiled texture Excel shape | scale texture fill C#
// Common Searches: Aspose.Cells set shape texture from web image | C# download image and use as shape fill in Excel | texture fill with tiling and scaling Aspose.Cells | fallback to local image when remote texture fails | apply byte array as texture fill Aspose.Cells
// Developer Intent: Load an image from a remote URL (or a local file if the download fails) and use it as a tiled, scaled texture fill for a specific shape in an Excel workbook generated with Aspose.Cells.
// Use Cases: Generate branded reports where a logo fetched from a web service fills a shape as a repeated pattern. | Create diagrammatic worksheets that apply externally hosted pattern images to shapes, automatically adjusting tile size. | Build templates that guarantee a texture fill by falling back to a bundled image when the online source is unavailable.
// AI Prompts: Write C# code that downloads an image from a URL, falls back to a local file, and applies it as a tiled, scaled texture fill to a rectangle shape using Aspose.Cells. | Show how to configure the TextureFill properties (IsTiling, Scale) for a shape after assigning image data from a byte array in Aspose.Cells. | Provide an example that applies different texture fills to multiple shapes, each using images from distinct URLs with error handling.

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a rectangle shape, downloads an image (with local fallback), sets the shape's FillType to Texture, enables tiling, scales the texture to 50 % and saves the file.
class Program
{
    static async Task Main(string[] args)
    {
        // URL of the image to be used as texture
        const string imageUrl = "https://example.com/texture.png";

        byte[] imageData = null;

        // Attempt to download the image; fallback to a local file if download fails
        try
        {
            using (HttpClient client = new HttpClient())
            {
                imageData = await client.GetByteArrayAsync(imageUrl);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to download image from '{imageUrl}'. Reason: {ex.Message}");
            const string localFallback = "defaultTexture.png";
            if (File.Exists(localFallback))
            {
                try
                {
                    imageData = File.ReadAllBytes(localFallback);
                    Console.WriteLine($"Loaded fallback texture from '{localFallback}'.");
                }
                catch (Exception fileEx)
                {
                    Console.WriteLine($"Failed to read fallback texture. Reason: {fileEx.Message}");
                }
            }
        }

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, height, width
        Shape rectangle = worksheet.Shapes.AddRectangle(2, 2, 5, 5, 150, 200);

        // Apply texture fill only if image data is available
        if (imageData != null && imageData.Length > 0)
        {
            rectangle.Fill.FillType = FillType.Texture;
            rectangle.Fill.TextureFill.ImageData = imageData;
            rectangle.Fill.TextureFill.IsTiling = true;
            rectangle.Fill.TextureFill.Scale = 0.5; // Scale the texture to 50%
        }
        else
        {
            Console.WriteLine("No texture applied to the shape because image data is unavailable.");
        }

        // Save the workbook to a file
        try
        {
            const string outputPath = "ShapeWithTexture.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook. Reason: {ex.Message}");
        }
    }
}
