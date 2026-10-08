// Title: Send a PNG snapshot of the first Excel worksheet to a Slack channel using Aspose.Cells and an incoming webhook in C#
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, renders the first worksheet to a PNG byte array, and uploads the image to a Slack incoming webhook as a file attachment. | Create an async .NET method that converts a worksheet to PNG, builds a multipart/form-data request with the image and a JSON text block, and posts it to a specified Slack webhook URL.
// Common Searches: how to render an Excel sheet to PNG with Aspose.Cells and send it to Slack in C# | C# upload worksheet image to Slack incoming webhook using multipart/form-data | Aspose.Cells export first worksheet as PNG and post to Slack channel | send Excel worksheet snapshot to Slack via webhook .NET example
// Tags: Aspose.Cells PNG rendering of worksheet | Slack webhook file attachment .NET | C# multipart/form-data image upload | Excel worksheet snapshot notification | async HTTP post to Slack webhook

using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an .xlsx workbook, uses Aspose.Cells to render the first worksheet as a PNG image in memory, then creates a multipart/form-data request that includes the image and a JSON text payload, and finally posts it to a Slack incoming webhook URL, reporting success or failure.
class Program
{
    // Replace with your actual Slack webhook URL
    private const string SlackWebhookUrl = "https://hooks.slack.com/services/XXXXX/XXXXX/XXXXXXXXXX";

    static async Task Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Set image rendering options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the worksheet to an image stored in a memory stream
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
            byte[] pngBytes;
            using (MemoryStream ms = new MemoryStream())
            {
                // Render the first page of the sheet (0‑based index)
                sheetRender.ToImage(0, ms);
                pngBytes = ms.ToArray();
            }

            // Prepare multipart/form-data content for Slack webhook
            using (var httpClient = new HttpClient())
            using (var content = new MultipartFormDataContent())
            {
                var imageContent = new ByteArrayContent(pngBytes);
                imageContent.Headers.ContentType = MediaTypeHeaderValue.Parse("image/png");
                // "file" is the name Slack expects for file uploads; adjust if necessary
                content.Add(imageContent, "file", "worksheet.png");

                // Slack incoming webhooks accept JSON payloads.
                var payload = new StringContent("{\"text\":\"Worksheet snapshot attached.\"}");
                payload.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");
                content.Add(payload, "payload_json");

                // Send the POST request to the Slack webhook URL
                HttpResponseMessage response = await httpClient.PostAsync(SlackWebhookUrl, content);
                if (response.IsSuccessStatusCode)
                {
                    Console.WriteLine("Image sent to Slack successfully.");
                }
                else
                {
                    Console.WriteLine($"Failed to send image. Status: {response.StatusCode}");
                    string responseBody = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"Response: {responseBody}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
