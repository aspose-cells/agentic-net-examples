// Title: Insert a Web Image into an Aspose.Cells Worksheet (C#)
// Description: This example creates a new Workbook, downloads an image from a given HTTP URL with HttpClient, streams it into a MemoryStream, and embeds the picture into the first worksheet at cell A1 using worksheet.Pictures.Add, then saves the file as an .xlsx workbook.
// Keywords: Aspose.Cells insert picture from URL | C# download image HttpClient | worksheet.Pictures.Add memory stream | embed web image Excel | Aspose.Cells add picture C#
// Common Searches: how to add an online image to Excel with Aspose.Cells | Aspose.Cells picture from HttpClient stream | C# embed JPEG from web address into worksheet | Aspose.Cells download and insert image
// Developer Intent: Embed an image retrieved from a remote URL directly into an Excel worksheet using Aspose.Cells.
// Use Cases: Automatically place a company logo hosted on a CDN into generated reports. | Build dashboards that pull icons or charts from external services and embed them in specific cells. | Create invoices that include branding assets stored on a web server without manual file handling.
// AI Prompts: Show how to set the picture's position and size after adding it with worksheet.Pictures.Add. | Generate code that iterates over a list of image URLs and adds each picture to a separate worksheet. | Explain how to verify the downloaded image format before embedding it in an Aspose.Cells workbook.

using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

namespace Example
{
    // This example creates a new Workbook, downloads an image from a given HTTP URL with HttpClient, streams it into a MemoryStream, and embeds the picture into the first worksheet at cell A1 using worksheet.Pictures.Add, then saves the file as an .xlsx workbook.
    class AddWebImageToWorksheet
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
                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = client.GetAsync(imageUrl).Result;
                    response.EnsureSuccessStatusCode();

                    using (Stream httpStream = response.Content.ReadAsStreamAsync().Result)
                    using (MemoryStream imageStream = new MemoryStream())
                    {
                        httpStream.CopyTo(imageStream);
                        imageStream.Position = 0; // reset position for reading

                        // Add the picture to the worksheet (row 1, column 1)
                        worksheet.Pictures.Add(1, 1, imageStream);
                    }
                }

                // Save the workbook
                string outputPath = "WorkbookWithWebImage.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"Error downloading image: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
