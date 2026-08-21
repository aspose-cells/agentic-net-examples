// Title: Convert Excel Worksheet to PNG and POST to REST API with Aspose.Cells (C#)
// Description: Creates a workbook, fills sample cells, renders the first worksheet to a PNG stream using Aspose.Cells SheetRender (OnePagePerSheet), and uploads the image to a REST endpoint with HttpClient, including success and error handling.
// Keywords: Aspose.Cells | C# | .NET | Excel to PNG | Worksheet image | SheetRender | ImageOrPrintOptions | HttpClient POST | REST API upload | image/png stream | asynchronous upload
// Common Searches: Aspose.Cells export worksheet as PNG C# | C# upload PNG image to web service | How to post an Excel sheet image with HttpClient | Convert Excel to image and call REST endpoint | Send worksheet snapshot to API .NET
// Developer Intent: Generate a PNG snapshot of an Excel worksheet and send it to a REST endpoint using C#.
// Use Cases: Capture a visual snapshot of a report sheet and deliver it to a reporting service for downstream processing. | Create a thumbnail of a spreadsheet for preview in a web portal and store it via a media‑storage API. | Transmit a worksheet image to an OCR or data‑extraction service that works with image inputs.
// AI Prompts: Write C# code that renders the first worksheet of a workbook to a PNG stream with Aspose.Cells and posts it to a given URL using HttpClient, handling exceptions and response status. | Show how to configure ImageOrPrintOptions for high‑resolution PNG output before uploading the image. | Demonstrate adding Bearer token authentication and custom headers to the HttpClient request when uploading the worksheet image. | Provide an example of retry logic for transient network failures during the PNG POST operation.

using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing.Imaging;

// Creates a workbook, fills sample cells, renders the first worksheet to a PNG stream using Aspose.Cells SheetRender (OnePagePerSheet), and uploads the image to a REST endpoint with HttpClient, including success and error handling.
class Program
{
    static async Task Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            using (Workbook workbook = new Workbook())
            {
                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["A1"].PutValue("Sample Data");
                worksheet.Cells["B1"].PutValue(123);
                worksheet.Cells["A2"].PutValue("More Data");
                worksheet.Cells["B2"].PutValue(456);

                // Configure image rendering options for PNG
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    // Default format is PNG; explicit setting removed to avoid compatibility issues
                    OnePagePerSheet = true // Render each sheet as a single page
                };

                // Render worksheet to an image stream
                using (MemoryStream imageStream = new MemoryStream())
                {
                    try
                    {
                        SheetRender sheetRender = new SheetRender(worksheet, options);
                        sheetRender.ToImage(0, imageStream);
                        imageStream.Position = 0;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error during rendering: {ex.Message}");
                        return;
                    }

                    // Prepare HTTP client
                    using (HttpClient httpClient = new HttpClient())
                    {
                        const string endpointUrl = "https://example.com/api/upload";

                        using (StreamContent httpContent = new StreamContent(imageStream))
                        {
                            httpContent.Headers.ContentType = new MediaTypeHeaderValue("image/png");

                            try
                            {
                                HttpResponseMessage response = await httpClient.PostAsync(endpointUrl, httpContent);
                                if (response.IsSuccessStatusCode)
                                {
                                    string responseBody = await response.Content.ReadAsStringAsync();
                                    Console.WriteLine("Upload successful. Server response:");
                                    Console.WriteLine(responseBody);
                                }
                                else
                                {
                                    Console.WriteLine($"Upload failed. Status code: {(int)response.StatusCode} {response.ReasonPhrase}");
                                }
                            }
                            catch (HttpRequestException ex)
                            {
                                Console.WriteLine($"HTTP request error: {ex.Message}");
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
