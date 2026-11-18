// Title: Insert a picture from an online URL into an Excel worksheet using Aspose.Cells for .NET
// Common Searches: asp.net download image and embed into Excel with Aspose.Cells | c# add picture from http url to worksheet cell using Aspose.Cells | how to use Pictures.Add with a MemoryStream in Aspose.Cells | save workbook with embedded web image using Aspose.Cells .NET | async download image and insert into Excel file programmatically
// Tags: Aspose.Cells Pictures.Add from MemoryStream | embed web image into Excel worksheet .NET | download image byte array for Excel shape | insert picture into cell A1 using Aspose.Cells | C# Aspose.Cells image embedding from URL

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The example creates a new Workbook, downloads an image from a specified web URL with HttpClient, wraps the bytes in a MemoryStream, and uses the worksheet's Pictures.Add method to place the picture in cell A1. The workbook is then saved as EmbeddedPicture.xlsx.
class Program
{
    static async Task Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // URL of the picture to embed
            var imageUrl = "https://example.com/image.png";

            // Download the image data
            byte[] imageBytes;
            using (var httpClient = new HttpClient())
            {
                imageBytes = await httpClient.GetByteArrayAsync(imageUrl);
            }

            // Ensure image data was retrieved
            if (imageBytes == null || imageBytes.Length == 0)
                throw new InvalidOperationException("Failed to download image data.");

            // Add the picture to cell A1 (row 0, column 0) using a MemoryStream
            using (var imageStream = new MemoryStream(imageBytes))
            {
                sheet.Pictures.Add(0, 0, imageStream);
            }

            // Save the workbook
            var outputPath = "EmbeddedPicture.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
