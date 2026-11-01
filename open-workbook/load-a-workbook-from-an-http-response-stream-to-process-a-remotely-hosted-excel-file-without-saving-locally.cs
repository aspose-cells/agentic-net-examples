// Title: Load an Excel workbook from an HTTP response stream with Aspose.Cells for .NET without saving to disk
// Common Searches: Aspose.Cells C# load workbook from HttpClient response stream | read remote .xlsx file directly into Aspose.Cells without creating a temporary file | process Excel file from URL using Aspose.Cells stream API in .NET | how to open an Excel workbook from a web request stream with Aspose.Cells | download Excel via HttpClient and pass stream to Aspose.Cells Workbook constructor
// Tags: Aspose.Cells stream workbook loading | C# HttpClient retrieve Excel stream | process remote .xlsx without temporary file | Aspose.Cells enumerate worksheet names | stream-based Excel handling with Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The example uses HttpClient to fetch an Excel file from a remote URL, reads the response content as a stream, and creates an Aspose.Cells Workbook directly from that stream. It then iterates through the workbook's worksheets to display their names, handling HTTP, I/O, and general exceptions without writing the file to disk.
class Program
{
    static async Task Main(string[] args)
    {
        // URL of the remote Excel file
        string fileUrl = "https://example.com/sample.xlsx";

        try
        {
            using HttpClient httpClient = new HttpClient();

            // Send GET request and obtain the response
            using HttpResponseMessage response = await httpClient.GetAsync(fileUrl, HttpCompletionOption.ResponseHeadersRead);

            if (!response.IsSuccessStatusCode)
            {
                Console.WriteLine($"Failed to download file. Status code: {(int)response.StatusCode} {response.ReasonPhrase}");
                return;
            }

            // Get the response content stream (no local file is created)
            using Stream excelStream = await response.Content.ReadAsStreamAsync();

            // Load the workbook directly from the stream
            Workbook workbook = new Workbook(excelStream);

            // Example processing: list all worksheet names
            Console.WriteLine("Worksheets in the workbook:");
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine("- " + sheet.Name);
            }

            // Additional processing can be performed here...
        }
        catch (HttpRequestException ex)
        {
            Console.WriteLine($"HTTP request error: {ex.Message}");
        }
        catch (IOException ex)
        {
            Console.WriteLine($"IO error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
