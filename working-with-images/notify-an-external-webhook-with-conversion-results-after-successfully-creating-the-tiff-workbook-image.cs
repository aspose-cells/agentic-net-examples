// Title: Convert an Aspose.Cells workbook to a TIFF file and send conversion details to an external webhook using C#
// AI Prompts: Write C# code that creates a workbook with Aspose.Cells, saves it as a TIFF image, and posts a JSON payload (file name, size, success flag, UTC timestamp) to a given webhook URL. | Add robust try‑catch logic so that if the TIFF export or the HTTP request fails, a failure JSON payload containing the error message is sent to the same webhook. | Demonstrate reusing a single HttpClient instance for both success and error webhook notifications after exporting a workbook to TIFF.
// Common Searches: how to export Aspose.Cells workbook to TIFF and notify a webhook in C# | c# Aspose.Cells save as tiff then post json to external service | send conversion result metadata to webhook after generating TIFF from Excel using Aspose.Cells | error handling for webhook notification after Aspose.Cells image export in .NET
// Tags: Aspose.Cells export workbook to TIFF C# | post JSON payload to webhook .NET | TIFF conversion result metadata Aspose.Cells | error handling webhook notification Aspose.Cells | HttpClient reuse for webhook calls C#

using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Aspose.Cells;

// Creates a simple workbook with Aspose.Cells, saves it as a TIFF image, gathers file information, and posts a JSON payload with the conversion result to a specified webhook. Includes error handling that sends a failure payload if the export or HTTP request fails.
class Program
{
    // Replace with your actual webhook URL
    private const string WebhookUrl = "https://example.com/webhook";

    static async Task Main(string[] args)
    {
        // Path for the generated TIFF image
        string tiffPath = "WorkbookImage.tiff";

        try
        {
            // ---------- Create a simple workbook ----------
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);

            // ---------- Save the workbook as a TIFF image ----------
            workbook.Save(tiffPath, SaveFormat.Tiff);

            // Verify that the file was created
            if (!File.Exists(tiffPath))
                throw new FileNotFoundException("TIFF file was not created.", tiffPath);

            // Gather conversion result details
            var fileInfo = new FileInfo(tiffPath);
            var resultPayload = new
            {
                FileName = fileInfo.Name,
                FileSizeBytes = fileInfo.Length,
                Success = true,
                CreatedAtUtc = DateTime.UtcNow
            };

            // Serialize payload to JSON
            string jsonPayload = JsonSerializer.Serialize(resultPayload);

            // ---------- Notify external webhook ----------
            using var httpClient = new HttpClient();
            var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
            HttpResponseMessage response = await httpClient.PostAsync(WebhookUrl, content);

            // Ensure the webhook responded successfully
            response.EnsureSuccessStatusCode();

            Console.WriteLine("Webhook notified successfully.");
        }
        catch (Exception ex)
        {
            // In case of any failure, send a failure notification to the webhook
            var errorPayload = new
            {
                FileName = Path.GetFileName(tiffPath),
                Success = false,
                ErrorMessage = ex.Message,
                CreatedAtUtc = DateTime.UtcNow
            };

            string errorJson = JsonSerializer.Serialize(errorPayload);
            using var httpClient = new HttpClient();
            var errorContent = new StringContent(errorJson, Encoding.UTF8, "application/json");
            try
            {
                await httpClient.PostAsync(WebhookUrl, errorContent);
            }
            catch
            {
                // Swallow any exceptions from the error notification to avoid recursive failures
            }

            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
