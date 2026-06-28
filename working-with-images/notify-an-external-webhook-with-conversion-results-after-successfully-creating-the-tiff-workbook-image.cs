using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
{
    static async Task Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Aspose.Cells TIFF Rendering Demo");

        // Configure image rendering options for TIFF output
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,
            TiffCompression = TiffCompression.CompressionLZW,
            HorizontalResolution = 300,
            VerticalResolution = 300,
            OnePagePerSheet = true
        };

        // Render the workbook to a TIFF file using the provided WorkbookRender.ToImage(string) method
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);
        string tiffPath = Path.Combine(outputDir, "workbook.tiff");
        WorkbookRender renderer = new WorkbookRender(workbook, options);
        renderer.ToImage(tiffPath);
        Console.WriteLine($"Workbook rendered successfully to: {tiffPath}");

        // Prepare payload for the external webhook
        var payload = new
        {
            success = true,
            filePath = tiffPath,
            fileSize = new FileInfo(tiffPath).Length,
            timestampUtc = DateTime.UtcNow
        };
        string jsonPayload = JsonSerializer.Serialize(payload);
        string webhookUrl = "https://example.com/webhook"; // replace with actual URL

        // Send POST request to the webhook
        using (HttpClient httpClient = new HttpClient())
        {
            var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
            try
            {
                HttpResponseMessage response = await httpClient.PostAsync(webhookUrl, content);
                Console.WriteLine($"Webhook response status: {response.StatusCode}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error notifying webhook: {ex.Message}");
            }
        }
    }
}