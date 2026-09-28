// Title: Render an Excel worksheet to a PNG image with Aspose.Cells and post it to a Slack channel via incoming webhook in C#
// AI Prompts: Create a C# console program that uses Aspose.Cells to convert the first worksheet of a workbook into a PNG stream and uploads the image to a Slack incoming webhook using HttpClient. | Add a JSON payload with a custom text caption while still attaching the PNG worksheet image to the Slack webhook request. | Switch the rendering options to generate a JPEG instead of PNG and adjust the multipart/form-data request for Slack accordingly.
// Common Searches: C# Aspose.Cells export first worksheet as PNG and send to Slack webhook | How to post an Excel sheet image to a Slack channel using HttpClient multipart request | Upload worksheet snapshot to Slack incoming webhook from a .NET console app | Aspose.Cells render worksheet to image and notify team on Slack | Send generated PNG from Excel to Slack via webhook in C#
// Tags: Aspose.Cells render worksheet to PNG | C# HttpClient multipart upload to Slack webhook | Excel sheet image notification Slack | Slack incoming webhook file attachment | MemoryStream image export Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

// The example creates a Workbook, fills cells with sample data, renders the first worksheet to a PNG image stored in a MemoryStream using Aspose.Cells, and then posts the image to a Slack incoming webhook via HttpClient with a multipart/form-data request.
class Program
{
    static async Task Main()
    {
        try
        {
            // 1. Create a workbook and fill it with sample data
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // 2. Render the first worksheet to a PNG image (in memory)
            var imgOptions = new ImageOrPrintOptions
            {
                // Default image format is PNG; no need to set explicitly
                OnePagePerSheet = true
            };
            var renderer = new SheetRender(sheet, imgOptions);
            using var pngStream = new MemoryStream();
            renderer.ToImage(0, pngStream);
            pngStream.Position = 0; // reset for reading

            // 3. Define the Slack incoming webhook URL (replace with real URL)
            const string slackWebhookUrl = "https://hooks.slack.com/services/XXXXX/XXXXX/XXXXX";

            // 4. Send the PNG image to Slack via HTTP POST (multipart/form-data)
            using var httpClient = new HttpClient();
            using var multipart = new MultipartFormDataContent();

            var imageContent = new ByteArrayContent(pngStream.ToArray());
            imageContent.Headers.ContentType = MediaTypeHeaderValue.Parse("image/png");
            multipart.Add(imageContent, "file", "worksheet.png");

            // Note: Incoming webhooks normally accept JSON payloads.
            // This example posts the image directly for demonstration purposes.
            HttpResponseMessage response = await httpClient.PostAsync(slackWebhookUrl, multipart);
            string responseText = await response.Content.ReadAsStringAsync();

            Console.WriteLine($"Slack response: {responseText}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
