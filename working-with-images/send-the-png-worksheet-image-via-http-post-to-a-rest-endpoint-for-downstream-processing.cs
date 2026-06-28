using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static async Task Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Aspose.Cells PNG Export");
        worksheet.Cells["A2"].PutValue(DateTime.Now);

        // Configure rendering options for PNG output
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            OnePagePerSheet = true
        };

        // Render the first page of the worksheet to a memory stream
        using (MemoryStream pngStream = new MemoryStream())
        {
            SheetRender sheetRender = new SheetRender(worksheet, renderOptions);
            sheetRender.ToImage(0, pngStream); // uses SheetRender.ToImage(int, Stream) rule
            pngStream.Position = 0; // reset stream for reading

            // Prepare HTTP client for POST
            using (HttpClient httpClient = new HttpClient())
            {
                string endpoint = "https://example.com/api/upload"; // replace with actual URL

                // Create HTTP content from the PNG stream
                ByteArrayContent httpContent = new ByteArrayContent(pngStream.ToArray());
                httpContent.Headers.ContentType = new MediaTypeHeaderValue("image/png");

                // Send POST request
                HttpResponseMessage response = await httpClient.PostAsync(endpoint, httpContent);
                string responseBody = await response.Content.ReadAsStringAsync();

                Console.WriteLine($"Response status: {response.StatusCode}");
                Console.WriteLine($"Response body: {responseBody}");
            }
        }
    }
}