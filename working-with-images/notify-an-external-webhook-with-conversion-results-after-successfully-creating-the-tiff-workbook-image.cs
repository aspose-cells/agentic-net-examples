// Title: Convert an Excel workbook to a multi-page TIFF with Aspose.Cells for .NET and send conversion details to an external webhook
// AI Prompts: Write C# code that creates a workbook, fills it with sample data, saves it as a multi-page TIFF using Aspose.Cells, and posts a JSON payload containing the file name, size, status, and timestamp to a given webhook URL. | Enhance the Aspose.Cells TIFF export sample with async HttpClient POST logic that includes retry handling and detailed error logging for the webhook call.
// Common Searches: how to export an Aspose.Cells workbook to a multi page TIFF and notify a webhook in C# | C# Aspose.Cells save workbook as TIFF then post file metadata to external API | async HttpClient POST after Aspose.Cells image conversion .NET | send JSON payload with file information to webhook after TIFF generation using Aspose.Cells
// Tags: Aspose.Cells export workbook to TIFF | C# post JSON payload with HttpClient | async webhook notification after file creation | multi-page TIFF generation using Aspose.Cells | retry and error handling for HTTP POST in .NET

using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Aspose.Cells;

// The program creates an Excel workbook, populates it with data, saves it as a multi-page TIFF using Aspose.Cells, verifies the file exists, and then asynchronously sends a JSON payload with the file name, size, status, and timestamp to a specified external webhook.
class Program
{
    static async Task Main(string[] args)
    {
        // Define the output TIFF file path
        string tiffPath = "output.tiff";

        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate the worksheet with sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(85);

            // Save the workbook as a TIFF image (each sheet becomes a page)
            workbook.Save(tiffPath, SaveFormat.Tiff);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during workbook creation or TIFF conversion: {ex.Message}");
            return;
        }

        // Verify that the TIFF file was created successfully
        if (File.Exists(tiffPath))
        {
            // Build the payload to send to the external webhook
            var payload = new
            {
                fileName = Path.GetFileName(tiffPath),
                fileSize = new FileInfo(tiffPath).Length,
                status = "Success",
                timestamp = DateTime.UtcNow
            };

            string jsonPayload = JsonSerializer.Serialize(payload);
            var httpContent = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

            // Replace with the actual webhook endpoint URL
            string webhookUrl = "https://example.com/webhook";

            using HttpClient httpClient = new HttpClient();

            try
            {
                HttpResponseMessage response = await httpClient.PostAsync(webhookUrl, httpContent);
                response.EnsureSuccessStatusCode();
                Console.WriteLine("Webhook notified successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to notify webhook: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("TIFF conversion failed; file not found.");
        }
    }
}
