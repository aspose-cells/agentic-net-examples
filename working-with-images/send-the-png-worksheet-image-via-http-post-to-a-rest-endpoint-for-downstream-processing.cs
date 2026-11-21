// Title: Render an Excel worksheet to PNG and upload it with HttpClient multipart/form-data in C# using Aspose.Cells
// Common Searches: how to render an Excel worksheet as PNG and send it with HttpClient in C# | Aspose.Cells convert worksheet to image stream for HTTP POST | upload Excel sheet image to REST API using multipart/form-data .NET | C# code to post worksheet PNG to web service endpoint | sending generated PNG from workbook to a web API with HttpClient
// Tags: Aspose.Cells worksheet PNG rendering | C# HttpClient multipart/form-data image upload | Excel sheet image POST to REST endpoint | convert worksheet to image stream Aspose.Cells | upload PNG to web API using .NET

using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing.Imaging;

// The example loads an Excel workbook, renders the first worksheet to a PNG image in memory with Aspose.Cells, then builds a multipart/form-data request containing the PNG and posts it to a specified REST endpoint using HttpClient, with error handling for file access, workbook loading, and HTTP failures.
class Program
{
    static async Task Main(string[] args)
    {
        // Path to the Excel file
        string workbookPath = "input.xlsx";

        // REST endpoint that will receive the PNG image
        string endpointUrl = "https://example.com/api/upload";

        try
        {
            await SendWorksheetImageAsync(workbookPath, endpointUrl);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    static async Task SendWorksheetImageAsync(string workbookPath, string endpointUrl)
    {
        // Verify that the workbook file exists
        if (!File.Exists(workbookPath))
            throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

        // Load the workbook from the specified file
        Workbook workbook;
        try
        {
            workbook = new Workbook(workbookPath);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to load workbook.", ex);
        }

        // Get the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Set rendering options (default image format is PNG)
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            OnePagePerSheet = true
        };

        // Render the worksheet to an image in memory
        SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
        using (MemoryStream imageStream = new MemoryStream())
        {
            // Render the first page (page index 0) of the worksheet
            sheetRender.ToImage(0, imageStream);
            imageStream.Position = 0;

            // Prepare an HTTP client for the POST request
            using (HttpClient httpClient = new HttpClient())
            {
                // Build multipart/form-data content with the PNG image
                using (MultipartFormDataContent multipart = new MultipartFormDataContent())
                {
                    ByteArrayContent imageContent = new ByteArrayContent(imageStream.ToArray());
                    imageContent.Headers.ContentType = MediaTypeHeaderValue.Parse("image/png");
                    multipart.Add(imageContent, "file", "worksheet.png");

                    // Send the POST request to the REST endpoint
                    HttpResponseMessage response;
                    try
                    {
                        response = await httpClient.PostAsync(endpointUrl, multipart);
                        response.EnsureSuccessStatusCode();
                    }
                    catch (Exception ex)
                    {
                        throw new HttpRequestException("Failed to upload image.", ex);
                    }
                }
            }
        }
    }
}
