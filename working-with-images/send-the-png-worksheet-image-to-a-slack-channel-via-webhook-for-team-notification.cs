using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static async Task Main()
    {
        // Slack incoming webhook URL (replace with your actual webhook)
        string webhookUrl = "https://hooks.slack.com/services/XXXXX/XXXXX/XXXXX";

        // -------------------------------------------------
        // Create a workbook and add some sample data
        // -------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Hello Slack!");
        worksheet.Cells["A2"].PutValue(DateTime.Now);

        // -------------------------------------------------
        // Render the first page of the worksheet to a PNG image in memory
        // -------------------------------------------------
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png,
            OnePagePerSheet = true
        };
        SheetRender sheetRender = new SheetRender(worksheet, renderOptions);

        using (MemoryStream imageStream = new MemoryStream())
        {
            sheetRender.ToImage(0, imageStream);   // Use provided SheetRender.ToImage overload
            imageStream.Position = 0;               // Reset stream position for reading

            // -------------------------------------------------
            // Prepare multipart/form-data payload for Slack webhook
            // -------------------------------------------------
            using (HttpClient httpClient = new HttpClient())
            using (MultipartFormDataContent multipart = new MultipartFormDataContent())
            {
                // Image part
                StreamContent imageContent = new StreamContent(imageStream);
                imageContent.Headers.ContentType = MediaTypeHeaderValue.Parse("image/png");
                multipart.Add(imageContent, "file", "worksheet.png");

                // JSON payload part (required by Slack for incoming webhooks)
                StringContent jsonPayload = new StringContent("{\"text\":\"Worksheet image attached\"}");
                jsonPayload.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data")
                {
                    Name = "payload_json"
                };
                multipart.Add(jsonPayload);

                // -------------------------------------------------
                // Send POST request to Slack webhook
                // -------------------------------------------------
                HttpResponseMessage response = await httpClient.PostAsync(webhookUrl, multipart);
                string responseBody = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"Slack response: {responseBody}");
            }
        }
    }
}