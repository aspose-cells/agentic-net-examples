// Title: Embed a picture from an online URL into an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Download an image from a web address with HttpClient, load it into a MemoryStream, and add it to a worksheet at a chosen cell using the Aspose.Cells C# API. | After inserting the picture, adjust its Width and Height properties in pixels via Aspose.Cells. | Save the workbook to a .xlsx file after embedding and resizing the image, then confirm the output path.
// Common Searches: how to add a picture from a URL to an Excel file using Aspose.Cells C# | Aspose.Cells insert image from web into worksheet memory stream | set picture width and height in pixels with Aspose.Cells API | download image with HttpClient and embed into Excel workbook Aspose.Cells | place picture at row 3 column 2 in Aspose.Cells worksheet
// Tags: insert picture from URL using Aspose.Cells memory stream | resize embedded image Aspose.Cells C# | download image with HttpClient for Excel workbook | add picture to specific cell coordinates Aspose.Cells | embed web image into .xlsx file Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, retrieves an image from a specified web URL using HttpClient, streams the image into the first worksheet at row 3, column 2, sets the picture's width to 200 px and height to 150 px, and saves the result as OutputWithPicture.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // URL of the picture to embed
            const string imageUrl = "https://example.com/image.png";

            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Download the image data from the web URL using HttpClient
            byte[] imageData;
            using (var httpClient = new HttpClient())
            {
                imageData = httpClient.GetByteArrayAsync(imageUrl).Result;
            }

            // Add the picture to the worksheet using a memory stream (embed directly)
            using (var ms = new MemoryStream(imageData))
            {
                // Parameters: upper-left row, upper-left column, picture stream
                int pictureIndex = sheet.Pictures.Add(2, 1, ms);
                // Adjust picture size or position if needed
                Picture picture = sheet.Pictures[pictureIndex];
                picture.Width = 200;   // width in pixels
                picture.Height = 150;  // height in pixels
            }

            // Save the workbook to a file
            const string outputPath = "OutputWithPicture.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
