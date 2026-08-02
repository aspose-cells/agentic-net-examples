// Title: Add a Web Image to an Excel Worksheet with Aspose.Cells for .NET
// Description: Downloads an image from a URL using HttpClient, inserts it into the first worksheet via worksheet.Pictures.Add, and saves the workbook with the embedded picture.
// Keywords: Aspose.Cells | C# | add picture from URL | embed image in Excel | worksheet.Pictures.Add | HttpClient image download | memory stream picture | Excel automation | logo from CDN | product catalog images
// Common Searches: Aspose.Cells insert image from URL | C# add picture to Excel from web | how to embed remote image in Excel using Aspose | worksheet.Pictures.Add example C# | download image and place in cell B2 Aspose.Cells
// Developer Intent: Insert a picture fetched from an online source directly into an Excel worksheet.
// Use Cases: Automatically add a company logo stored on a CDN to generated reports. | Build product catalogs where each item’s picture is pulled from a remote server and positioned in the appropriate row. | Create invoices that embed customer signature images retrieved from a web service.
// AI Prompts: Generate C# code that uses Aspose.Cells to download an image from a given URL and embed it at cell B2, with proper error handling and resource disposal. | Show how to loop through a list of image URLs and add each picture to different cells in a worksheet using Aspose.Cells and memory streams. | Write a reusable method for Aspose.Cells that accepts a worksheet, image URL, row, and column, then downloads and inserts the picture.

using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Downloads an image from a URL using HttpClient, inserts it into the first worksheet via worksheet.Pictures.Add, and saves the workbook with the embedded picture.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // URL of the image to embed
            string imageUrl = "https://example.com/sample.jpg";

            // Download the image into a memory stream
            using (HttpClient httpClient = new HttpClient())
            {
                Stream imageStream = null;
                try
                {
                    // Get response and ensure success
                    HttpResponseMessage response = httpClient.GetAsync(imageUrl).Result;
                    response.EnsureSuccessStatusCode();
                    imageStream = response.Content.ReadAsStreamAsync().Result;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to download image: {ex.Message}");
                }

                if (imageStream != null)
                {
                    // Add the picture to the worksheet at row 1, column 1 (zero‑based indices)
                    worksheet.Pictures.Add(1, 1, imageStream);
                    imageStream.Dispose();
                }
                else
                {
                    Console.WriteLine("Image stream is null; picture will not be added.");
                }
            }

            // Save the workbook with the embedded picture
            string outputPath = "output_with_embedded_picture.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
