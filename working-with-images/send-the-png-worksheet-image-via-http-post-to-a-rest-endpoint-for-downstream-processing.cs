// Title: Render an Excel worksheet to PNG and upload it via HTTP POST using Aspose.Cells in C#
// Common Searches: C# generate PNG from first worksheet using Aspose.Cells and send to web API | Upload Excel worksheet image to REST service with HttpClient multipart/form-data | Aspose.Cells render worksheet to memory stream and post as image/png | Send worksheet screenshot to external endpoint from .NET console app | Convert Excel sheet to PNG and call POST endpoint in C#
// Tags: Aspose.Cells PNG rendering of worksheet | C# HttpClient form-data image POST | SheetRender ToImage memory stream usage | Upload PNG to REST API from .NET | Excel sheet image conversion with Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The console program loads a workbook, renders the first worksheet to a PNG image in a memory stream using Aspose.Cells' SheetRender, and then posts the image as multipart/form-data to a specified REST endpoint with HttpClient, handling responses and errors.
class Program
{
    static async Task Main(string[] args)
    {
        // Expect two arguments: path to the workbook and the REST endpoint URL
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: Program <workbookPath> <endpointUrl>");
            return;
        }

        string workbookPath = args[0];
        string endpointUrl = args[1];

        await SendWorksheetImageAsync(workbookPath, endpointUrl);
    }

    static async Task SendWorksheetImageAsync(string workbookPath, string endpointUrl)
    {
        // Verify that the workbook file exists
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Workbook file not found: {workbookPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(workbookPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            // Select the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Set rendering options (default image format is PNG)
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the worksheet to a memory stream (no temporary file needed)
            using (MemoryStream imageStream = new MemoryStream())
            {
                SheetRender renderer = new SheetRender(sheet, options);
                renderer.ToImage(0, imageStream);
                imageStream.Position = 0; // Reset stream position for reading

                // Prepare HTTP client and multipart/form-data content
                using (HttpClient client = new HttpClient())
                using (MultipartFormDataContent multipart = new MultipartFormDataContent())
                {
                    StreamContent fileContent = new StreamContent(imageStream);
                    fileContent.Headers.ContentType = new MediaTypeHeaderValue("image/png");
                    multipart.Add(fileContent, "file", "worksheet.png");

                    try
                    {
                        // Send the PNG image via HTTP POST to the REST endpoint
                        HttpResponseMessage response = await client.PostAsync(endpointUrl, multipart);
                        response.EnsureSuccessStatusCode();

                        // Output server response for verification
                        string responseBody = await response.Content.ReadAsStringAsync();
                        Console.WriteLine($"Server response: {responseBody}");
                    }
                    catch (Exception httpEx)
                    {
                        Console.WriteLine($"HTTP request failed: {httpEx.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
