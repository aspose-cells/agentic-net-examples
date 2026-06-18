using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    // Replace with your actual Slack Incoming Webhook URL
    private const string SlackWebhookUrl = "https://hooks.slack.com/services/XXXXX/XXXXX/XXXXXXXXXXXXXXXXXXXX";

    static async Task Main()
    {
        // 1. Create a sample workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Aspose.Cells worksheet rendered to PNG");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // 2. Render the first page of the worksheet to a PNG image in memory
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png,
            OnePagePerSheet = true
        };

        SheetRender renderer = new SheetRender(sheet, renderOptions);
        using (MemoryStream imageStream = new MemoryStream())
        {
            // Render page 0 to the memory stream
            renderer.ToImage(0, imageStream);
            renderer.Dispose();

            // 3. Convert the image bytes to a Base64 string (data URI)
            string base64Image = Convert.ToBase64String(imageStream.ToArray());
            string dataUri = $"data:image/png;base64,{base64Image}";

            // 4. Build the JSON payload for Slack
            string payloadJson = $@"{{
                ""text"": ""Worksheet image preview:"",
                ""attachments"": [
                    {{
                        ""fallback"": ""Worksheet image"",
                        ""image_url"": ""{dataUri}"",
                        ""title"": ""Worksheet.png""
                    }}
                ]
            }}";

            // 5. Send the payload to Slack via HTTP POST
            using (HttpClient httpClient = new HttpClient())
            {
                var content = new StringContent(payloadJson, Encoding.UTF8, "application/json");
                HttpResponseMessage response = await httpClient.PostAsync(SlackWebhookUrl, content);
                string responseBody = await response.Content.ReadAsStringAsync();

                Console.WriteLine($"Slack response: {(int)response.StatusCode} {responseBody}");
            }
        }
    }
}