using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static async System.Threading.Tasks.Task Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Sample Data for PNG rendering");

        // Configure image rendering options (PNG)
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            OnePagePerSheet = true
        };

        // Render the first page of the worksheet to a memory stream (SheetRender.ToImage(int, Stream) rule)
        using (MemoryStream imageStream = new MemoryStream())
        {
            SheetRender sheetRender = new SheetRender(worksheet, options);
            sheetRender.ToImage(0, imageStream);
            imageStream.Position = 0; // reset for reading

            // Prepare HTTP client for POST
            using (HttpClient client = new HttpClient())
            {
                string endpoint = "https://example.com/api/upload"; // replace with actual URL

                // Create content from the image stream
                var content = new StreamContent(imageStream);
                content.Headers.ContentType = new MediaTypeHeaderValue("image/png");

                // Send POST request
                HttpResponseMessage response = await client.PostAsync(endpoint, content);
                Console.WriteLine($"POST response: {(int)response.StatusCode} {response.ReasonPhrase}");
            }
        }
    }
}