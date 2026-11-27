// Title: Generate a PNG image of the first worksheet with Aspose.Cells and send it to a Slack channel via Incoming Webhook in C#
// Common Searches: C# how to convert an Excel worksheet to PNG using Aspose.Cells and post to Slack webhook | Aspose.Cells render first sheet as PNG and send image to Slack incoming webhook with HttpClient | Send base64 encoded worksheet image to a Slack channel from a .NET application | Create Slack notification with an Excel sheet image using Aspose.Cells in C# | Render Excel to PNG in memory and post to Slack via webhook using C#
// Tags: Aspose.Cells render worksheet to PNG C# | Slack incoming webhook image payload .NET | convert Excel sheet to base64 image C# | HttpClient post JSON to Slack webhook | in-memory PNG generation Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSlackNotifier
{
    // This C# example loads an Excel workbook with Aspose.Cells, renders the first worksheet to a PNG image in memory, encodes the image as a Base64 data URI, builds a Slack incoming webhook JSON payload containing an image block, and posts it to a Slack channel using HttpClient while handling success and error responses.
    class Program
    {
        // Entry point – async to allow awaiting HTTP calls.
        static async Task Main(string[] args)
        {
            // -----------------------------------------------------------------
            // Configuration – adjust these paths or set via environment variables.
            // -----------------------------------------------------------------
            const string inputExcelPath = "input.xlsx"; // Path to the source workbook.
            const string slackWebhookEnv = "SLACK_WEBHOOK_URL"; // Environment variable name.

            // Retrieve the Slack webhook URL from environment.
            string slackWebhookUrl = Environment.GetEnvironmentVariable(slackWebhookEnv);
            if (string.IsNullOrWhiteSpace(slackWebhookUrl))
            {
                Console.Error.WriteLine($"Error: Environment variable '{slackWebhookEnv}' is not set.");
                return;
            }

            // Verify the input Excel file exists.
            if (!File.Exists(inputExcelPath))
            {
                Console.Error.WriteLine($"Error: Input Excel file '{inputExcelPath}' not found.");
                return;
            }

            try
            {
                // -------------------------------------------------------------
                // Load workbook and render the first worksheet page to PNG.
                // -------------------------------------------------------------
                Workbook workbook = new Workbook(inputExcelPath);
                Worksheet sheet = workbook.Worksheets[0]; // Use the first worksheet.

                // Configure rendering options.
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    OnePagePerSheet = true,
                    HorizontalResolution = 150,
                    VerticalResolution = 150
                };

                // Create the renderer.
                SheetRender renderer = new SheetRender(sheet, renderOptions);

                // Render the first (and only) page to a memory stream.
                using (MemoryStream pngStream = new MemoryStream())
                {
                    renderer.ToImage(0, pngStream);
                    pngStream.Position = 0; // Reset for reading.

                    // Convert the PNG bytes to a Base64 string for inclusion in JSON.
                    string base64Image = Convert.ToBase64String(pngStream.ToArray());

                    // ---------------------------------------------------------
                    // Build the JSON payload expected by Slack Incoming Webhook.
                    // Slack does not natively accept raw image bytes; however,
                    // many integrations accept a Base64 string inside a custom
                    // field. This example includes the image as a block with
                    // a data URI so that Slack can render it if supported.
                    // ---------------------------------------------------------
                    string dataUri = $"data:image/png;base64,{base64Image}";
                    var payload = new
                    {
                        text = "📊 Worksheet rendered as PNG:",
                        blocks = new object[]
                        {
                            new
                            {
                                type = "section",
                                text = new
                                {
                                    type = "mrkdwn",
                                    text = "Here is the rendered worksheet image."
                                }
                            },
                            new
                            {
                                type = "image",
                                image_url = dataUri,
                                alt_text = "Worksheet PNG"
                            }
                        }
                    };

                    string jsonPayload = System.Text.Json.JsonSerializer.Serialize(payload);

                    // ---------------------------------------------------------
                    // Send the payload to Slack via HTTP POST.
                    // ---------------------------------------------------------
                    using (HttpClient httpClient = new HttpClient())
                    {
                        var httpContent = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
                        HttpResponseMessage response = await httpClient.PostAsync(slackWebhookUrl, httpContent);

                        if (response.IsSuccessStatusCode)
                        {
                            Console.WriteLine("Success: Image posted to Slack channel.");
                        }
                        else
                        {
                            string responseBody = await response.Content.ReadAsStringAsync();
                            Console.Error.WriteLine($"Error: Slack webhook returned {(int)response.StatusCode} - {response.ReasonPhrase}");
                            Console.Error.WriteLine($"Response body: {responseBody}");
                        }
                    }
                }

                // Clean up Aspose objects.
                renderer.Dispose();
                workbook.Dispose();
            }
            catch (Exception ex)
            {
                // Any unexpected exception is reported to the console.
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
