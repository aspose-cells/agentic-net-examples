// Title: Render a Workbook to LZW‑Compressed TIFF and Send Conversion Details to a Webhook (C# Aspose.Cells)
// Description: This example shows how to create an Excel workbook with Aspose.Cells, configure ImageOrPrintOptions for a 300 dpi LZW‑compressed multi‑page TIFF, render the workbook to a MemoryStream using WorkbookRender.ToImage, build a JSON payload (success, format, sizeBytes, timestamp), post it to an external webhook via HttpClient, and optionally save the TIFF file locally.
// Keywords: Aspose.Cells TIFF rendering C# | WorkbookRender ToImage MemoryStream | LZW compression 300 dpi TIFF | C# post JSON webhook | Excel to TIFF conversion metadata | Aspose.Cells webhook notification | save TIFF from stream | Aspose.Cells image export
// Common Searches: How to export an Aspose.Cells workbook to a TIFF image in C# | C# Aspose.Cells render workbook as multi‑page TIFF | Send conversion results to a webhook after generating TIFF with Aspose.Cells | Post JSON payload from C# after Aspose.Cells image rendering | Save TIFF from Aspose.Cells MemoryStream
// Developer Intent: Generate a high‑resolution TIFF from an Excel workbook with Aspose.Cells and automatically report the conversion outcome to an external webhook.
// Use Cases: Automated document pipelines that convert Excel files to searchable TIFFs and notify a document‑management service. | Batch processing jobs that need to log file size, format, and timestamp for compliance auditing. | CI/CD workflows where conversion status is sent to a monitoring endpoint before proceeding to the next step.
// AI Prompts: Create C# code that renders an Aspose.Cells workbook to a 300 dpi LZW‑compressed multi‑page TIFF and posts a JSON payload with success, format, sizeBytes, and timestamp to a given webhook URL. | Add retry logic and detailed error handling to the webhook call after TIFF rendering, including logging of HTTP response codes. | Refactor the sample to stream the TIFF directly to the webhook without writing a local file, while still returning conversion metadata.

using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsWebhookDemo
{
    // This example shows how to create an Excel workbook with Aspose.Cells, configure ImageOrPrintOptions for a 300 dpi LZW‑compressed multi‑page TIFF, render the workbook to a MemoryStream using WorkbookRender.ToImage, build a JSON payload (success, format, sizeBytes, timestamp), post it to an external webhook via HttpClient, and optionally save the TIFF file locally.
    public class Program
    {
        // Entry point
        public static async Task Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells TIFF Rendering Demo");

            // Configure image rendering options for TIFF output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,                     // Render as TIFF
                TiffCompression = TiffCompression.CompressionLZW,
                HorizontalResolution = 300,
                VerticalResolution = 300,
                OnePagePerSheet = true
            };

            // Render the whole workbook to a memory stream using the provided ToImage method
            using (MemoryStream tiffStream = new MemoryStream())
            {
                WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);
                renderer.ToImage(tiffStream); // Rule: WorkbookRender.ToImage(Stream)

                // Prepare webhook notification payload
                var payload = new
                {
                    success = true,
                    format = "tiff",
                    sizeBytes = tiffStream.Length,
                    timestamp = DateTime.UtcNow
                };
                string json = System.Text.Json.JsonSerializer.Serialize(payload);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                // Send POST request to external webhook
                using (HttpClient httpClient = new HttpClient())
                {
                    // Replace with your actual webhook URL
                    string webhookUrl = "https://example.com/webhook";

                    HttpResponseMessage response = await httpClient.PostAsync(webhookUrl, content);
                    if (response.IsSuccessStatusCode)
                    {
                        Console.WriteLine("Webhook notified successfully.");
                    }
                    else
                    {
                        Console.WriteLine($"Webhook notification failed. Status: {response.StatusCode}");
                    }
                }

                // Optionally, save the TIFF to a file for verification
                string outputPath = Path.Combine("output", "workbook_render.tiff");
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));
                File.WriteAllBytes(outputPath, tiffStream.ToArray());
                Console.WriteLine($"TIFF image saved to: {outputPath}");
            }
        }
    }
}
