// Title: Load an Excel workbook from an HTTP response stream with Aspose.Cells for .NET without writing to disk
// AI Prompts: Write C# code that uses HttpClient to fetch a .xlsx file and opens it directly with Aspose.Cells Workbook from the response stream. | Demonstrate how to modify a remotely downloaded Excel workbook in memory and export the result to a MemoryStream using Aspose.Cells. | Show error handling for HTTP failures when loading an Excel file into Aspose.Cells without creating a temporary file.
// Common Searches: asp.net core load remote xlsx into Aspose.Cells workbook from HttpClient stream | c# read excel file from http response without saving to disk | process downloaded Excel in memory using Aspose.Cells SaveFormat Xlsx | how to use Aspose.Cells to open an Excel file from a network stream | Aspose.Cells load workbook from Stream example c#
// Tags: load workbook from HttpClient response stream | process remote xlsx in memory Aspose.Cells | save Aspose.Cells workbook to MemoryStream | download excel with HttpClient without file system | Aspose.Cells open workbook from Stream C#

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The example downloads an .xlsx file via HttpClient, loads it directly into an Aspose.Cells Workbook from the response stream, prints the first worksheet name, and shows how to save the modified workbook to a MemoryStream without touching the file system.
class Program
{
    static async Task Main()
    {
        // URL of the remote Excel file
        const string url = "https://example.com/sample.xlsx";

        try
        {
            using HttpClient client = new HttpClient();

            // Send request and verify success status
            HttpResponseMessage response = await client.GetAsync(url);
            if (!response.IsSuccessStatusCode)
            {
                Console.WriteLine($"Failed to download file. HTTP {(int)response.StatusCode} {response.ReasonPhrase}");
                return;
            }

            // Read the content as a stream without saving to disk
            await using Stream responseStream = await response.Content.ReadAsStreamAsync();

            // Load the workbook directly from the response stream
            Workbook workbook = new Workbook(responseStream);

            // Example processing: output the name of the first worksheet
            Worksheet firstSheet = workbook.Worksheets[0];
            Console.WriteLine($"First worksheet name: {firstSheet.Name}");

            // Additional processing can be performed here
            // ...

            // If you need to obtain the modified workbook as bytes without writing to a file,
            // save it to a MemoryStream
            await using MemoryStream memoryStream = new MemoryStream();
            workbook.Save(memoryStream, SaveFormat.Xlsx);
            // memoryStream now contains the Excel file data
        }
        catch (HttpRequestException ex)
        {
            Console.WriteLine($"HTTP request error: {ex.Message}");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File not found: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
