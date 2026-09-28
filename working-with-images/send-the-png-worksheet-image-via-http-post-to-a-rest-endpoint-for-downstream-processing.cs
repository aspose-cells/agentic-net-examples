// Title: Render an Excel worksheet to a PNG image and upload it with HttpClient multipart POST in C# using Aspose.Cells
// AI Prompts: Generate C# code that uses Aspose.Cells to render the first worksheet of a workbook to a PNG stream and sends it to a given REST endpoint with HttpClient multipart/form-data. | Show how to add bearer‑token authentication to the HttpClient request when uploading the PNG image generated from an Aspose.Cells worksheet. | Provide an example that batches multiple worksheet PNG streams into a single multipart/form-data POST request using Aspose.Cells and HttpClient.
// Common Searches: how to export an Excel sheet as PNG and post it to a web API using Aspose.Cells | C# Aspose.Cells render worksheet to image and upload with HttpClient multipart | sending Excel worksheet image to REST service with authentication in .NET | convert Excel worksheet to PNG in memory and call external API from C# | Aspose.Cells SheetRender PNG upload example
// Tags: export worksheet to PNG with Aspose.Cells | multipart/form-data image upload via HttpClient | in‑memory PNG stream from Excel sheet | authenticated POST of PNG to web service | Aspose.Cells SheetRender PNG generation | C# send workbook image to API

using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Drawing.Imaging;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;   // Required for ImageOrPrintOptions and SheetRender

// The program creates a workbook, renders the first worksheet to a PNG image in a memory stream, and uploads the image to a REST API using HttpClient with multipart/form-data, optionally supporting authentication.
class Program
{
    static async Task Main(string[] args)
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["B2"].PutValue(12345);

            // Configure image export options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the worksheet to an image and store it in a memory stream
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            using (MemoryStream imageStream = new MemoryStream())
            {
                // Export the first page (index 0) of the worksheet as PNG
                renderer.ToImage(0, imageStream);
                imageStream.Position = 0; // Reset stream position for reading

                // Prepare HTTP client for POST request
                using (HttpClient httpClient = new HttpClient())
                {
                    // Create multipart/form-data content
                    using (MultipartFormDataContent multipartContent = new MultipartFormDataContent())
                    {
                        // Add the PNG image as a byte array content
                        ByteArrayContent imageContent = new ByteArrayContent(imageStream.ToArray());
                        imageContent.Headers.ContentType = new MediaTypeHeaderValue("image/png");
                        multipartContent.Add(imageContent, "file", "worksheet.png");

                        // Define the REST endpoint URL
                        string endpointUrl = "https://example.com/api/upload";

                        // Send POST request
                        HttpResponseMessage response = await httpClient.PostAsync(endpointUrl, multipartContent);
                        response.EnsureSuccessStatusCode();

                        // Optionally read response content
                        string responseBody = await response.Content.ReadAsStringAsync();
                        Console.WriteLine("Upload successful. Server response:");
                        Console.WriteLine(responseBody);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
