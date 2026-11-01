// Title: Load an Excel workbook from a remote URL with HttpClient, fallback to a local file, and read cell A1 using Aspose.Cells for .NET
// AI Prompts: Create C# code that fetches an .xlsx file from a given URL using HttpClient, loads it into an Aspose.Cells Workbook from the response stream, and prints the value of cell A1 on the first worksheet. | Add logic so that if the HTTP download fails, the program loads the workbook from a specified local .xlsx path, handling HttpRequestException and file‑access exceptions. | Enhance the sample with detailed try‑catch blocks that log errors for network failures, invalid workbook streams, and worksheet processing problems.
// Common Searches: asp.net download excel file with HttpClient and open it using Aspose.Cells | c# load workbook from stream fallback to local file Aspose.Cells | read cell value from first worksheet after downloading xlsx in .NET | handle HttpRequestException when loading Excel workbook with Aspose.Cells | example code for loading Excel from URL and reading A1 using Aspose.Cells
// Tags: httpclient fetch xlsx Aspose.Cells | stream based workbook loading Aspose.Cells C# | local fallback workbook Aspose.Cells | retrieve A1 cell value Aspose.Cells | robust workbook loading error handling .NET

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The example demonstrates how to download an .xlsx file from a remote URL using HttpClient, load it into an Aspose.Cells Workbook via a stream, and read the value of cell A1 from the first worksheet. It includes a fallback to a local file when the download fails and provides comprehensive exception handling for network, file, and processing errors.
class Program
{
    static async Task Main()
    {
        string url = "https://example.com/sample.xlsx";
        string localPath = "sample.xlsx";

        Workbook workbook = null;

        // Attempt to download the workbook from the URL
        try
        {
            using (HttpClient httpClient = new HttpClient())
            using (Stream stream = await httpClient.GetStreamAsync(url))
            {
                workbook = new Workbook(stream);
            }
        }
        catch (HttpRequestException ex)
        {
            Console.WriteLine($"Failed to download file: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error during download: {ex.Message}");
        }

        // Fallback to local file if download failed
        if (workbook == null)
        {
            if (File.Exists(localPath))
            {
                try
                {
                    workbook = new Workbook(localPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load local workbook: {ex.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Local file '{localPath}' not found.");
                return;
            }
        }

        // Process the workbook
        try
        {
            Worksheet firstSheet = workbook.Worksheets[0];
            Cell cellA1 = firstSheet.Cells["A1"];
            Console.WriteLine($"Value in A1: {cellA1.Value}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
