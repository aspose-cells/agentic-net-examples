// Title: Embed a picture from an online URL into an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that downloads an image from a given URL using HttpClient and inserts it into a worksheet at cell A1 with Aspose.Cells, using a MemoryStream. | Show how to resize an embedded picture (set width and height in pixels) after adding it to an Aspose.Cells worksheet. | Provide a complete example that saves the workbook after embedding the image and prints the full output path. | Explain how to handle exceptions when downloading the image and adding it to the worksheet.
// Common Searches: aspnet embed image from http url into excel using aspose.cells | c# download picture with HttpClient and add to aspose.cells worksheet | set picture width height after adding to aspose.cells workbook | embed image as embedded object not linked in excel using aspose.cells | asp.net core save excel with embedded picture from web
// Tags: Aspose.Cells Pictures.Add MemoryStream usage | embed online image into Excel workbook C# | HttpClient image retrieval for Aspose.Cells | adjust picture size Aspose.Cells | save Excel file with embedded picture Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, downloads an image from a specified web URL using HttpClient, embeds the image into the first worksheet at cell A1 via a MemoryStream, optionally sets the picture's width and height, and saves the workbook as EmbeddedPicture.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // URL of the picture to embed
            string imageUrl = "https://example.com/image.png";

            // Download the image data using HttpClient
            byte[] imageData;
            using (HttpClient client = new HttpClient())
            {
                imageData = client.GetByteArrayAsync(imageUrl).Result;
            }

            // Add the picture to the worksheet using a memory stream (embedded, not linked)
            using (MemoryStream ms = new MemoryStream(imageData))
            {
                // Add picture at cell A1 (row 0, column 0). The Add method returns the picture index.
                int pictureIndex = sheet.Pictures.Add(0, 0, ms);

                // Retrieve the Picture object using the returned index
                Picture picture = sheet.Pictures[pictureIndex];

                // Optional: set size of the picture (pixels)
                picture.Width = 200;
                picture.Height = 150;
            }

            // Save the workbook to a file
            string outputPath = "EmbeddedPicture.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
