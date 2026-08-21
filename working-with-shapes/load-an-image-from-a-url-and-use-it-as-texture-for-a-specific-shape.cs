// Title: Apply a Web Image as Texture Fill to a Shape in Aspose.Cells for .NET
// Description: Demonstrates how to download an image from a URL with HttpClient, assign it to a shape's TextureFill, enable tiling, adjust scaling, and save the workbook as an Excel file using Aspose.Cells.
// Keywords: Aspose.Cells texture fill | shape fill image .NET | download image for Excel shape | texture tiling Aspose.Cells | TextureFill.Scale property | C# Aspose.Cells shape example
// Common Searches: Aspose.Cells use online image as shape texture | C# set rectangle fill to downloaded picture in Excel | how to enable tiling for shape texture in Aspose.Cells | apply web image to shape fill Aspose.Cells .NET | scale texture fill for Excel shape programmatically
// Developer Intent: Load an image from a web address and apply it as a texture fill to a specific worksheet shape.
// Use Cases: Create a rectangle and fill it with a JPEG retrieved from a public URL. | Repeat the texture across the shape by turning on tiling. | Resize the image inside the shape using the Scale property (e.g., 0.5 for 50%). | Provide fallback logic when the image download fails.
// AI Prompts: Generate C# code that downloads an image from a URL and sets it as a TextureFill for a shape in Aspose.Cells, with tiling enabled. | Show how to adjust TextureFill.Scale to 0.75 and handle HttpRequestException in Aspose.Cells. | Create a reusable method that takes an image URL and a Shape object, applies the image as a texture fill, and returns the updated Workbook.

using System;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to download an image from a URL with HttpClient, assign it to a shape's TextureFill, enable tiling, adjust scaling, and save the workbook as an Excel file using Aspose.Cells.
class Program
{
    static async Task Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape that will receive the texture fill
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 200, 100);

            // Set the fill type of the shape to Texture
            shape.Fill.FillType = FillType.Texture;

            // Obtain the TextureFill object to assign image data
            TextureFill textureFill = shape.Fill.TextureFill;

            // URL of the image to be used as texture (valid image URL)
            string imageUrl = "https://www.gstatic.com/webp/gallery/1.jpg";

            // Download the image bytes
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    byte[] imageBytes = await client.GetByteArrayAsync(imageUrl);
                    // Assign the downloaded image data to the texture fill
                    textureFill.ImageData = imageBytes;
                }
                catch (HttpRequestException ex)
                {
                    Console.WriteLine($"Failed to download image: {ex.Message}");
                    // Optionally, load a local fallback image if needed
                }
            }

            // Optionally enable tiling or adjust scaling
            textureFill.IsTiling = true;
            textureFill.Scale = 0.5; // 50% scale

            // Save the workbook to a file
            workbook.Save("ShapeWithTexture.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
