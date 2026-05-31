using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // URL of the image to be used as texture
            const string imageUrl = "https://example.com/texture.png";

            // Download the image data from the URL (fallback to local file if needed)
            byte[] imageData = GetImageData(imageUrl);

            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
            Shape shape = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 200, 200);

            if (imageData != null)
            {
                // Set the fill type of the shape to texture and assign the image
                shape.Fill.FillType = FillType.Texture;
                shape.Fill.TextureFill.ImageData = imageData;
                shape.Fill.TextureFill.IsTiling = true; // optional tiling
            }
            else
            {
                Console.WriteLine("Image data not available; shape will use default fill.");
            }

            // Save the workbook with the textured shape
            workbook.Save("ShapeWithTexture.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Retrieves image bytes from a URL or a local fallback file
    private static byte[] GetImageData(string url)
    {
        try
        {
            return DownloadImageAsync(url).GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to download image from URL: {ex.Message}");
        }

        // Fallback to a local file named "texture.png" if it exists
        const string localPath = "texture.png";
        if (File.Exists(localPath))
        {
            try
            {
                return File.ReadAllBytes(localPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read local image file: {ex.Message}");
            }
        }

        // No image data available
        return null;
    }

    // Helper method to download image bytes from a URL
    private static async Task<byte[]> DownloadImageAsync(string url)
    {
        using var client = new HttpClient();
        return await client.GetByteArrayAsync(url);
    }
}