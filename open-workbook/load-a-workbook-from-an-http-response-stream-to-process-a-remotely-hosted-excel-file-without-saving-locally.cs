// Title: Read an Excel workbook from a remote URL using Aspose.Cells for .NET via HttpClient stream without saving to disk
// Common Searches: Aspose.Cells C# load workbook from HttpClient response stream | How to open an Excel file from a URL in memory using Aspose.Cells | Read cell values from a remote .xlsx file without writing to disk in .NET | List worksheet names from a streamed Excel workbook with Aspose.Cells | Download Excel file with HttpClient and pass stream to Aspose.Cells Workbook constructor
// Tags: load workbook directly from HttpClient stream | process remote xlsx in memory with Aspose.Cells | enumerate worksheets in streamed Excel workbook | read cell A1 from in‑memory workbook | download Excel file without local storage in .NET

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The example uses HttpClient to fetch an Excel file from a URL, obtains the response as a stream, and creates an Aspose.Cells Workbook directly from that stream. It then enumerates all worksheet names and reads the value of cell A1 from the first sheet, performing all operations in memory without writing any files to disk, with comprehensive error handling.
class Program
{
    static async Task Main(string[] args)
    {
        // URL of the remote Excel file
        string fileUrl = "https://example.com/sample.xlsx";

        try
        {
            // Create an HttpClient to download the file
            using (HttpClient httpClient = new HttpClient())
            {
                // Send GET request and obtain the response stream
                using (HttpResponseMessage response = await httpClient.GetAsync(fileUrl, HttpCompletionOption.ResponseHeadersRead))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine($"Failed to download file. HTTP status: {response.StatusCode}");
                        return;
                    }

                    // Get the response content stream (no local file is created)
                    using (Stream responseStream = await response.Content.ReadAsStreamAsync())
                    {
                        // Load the workbook directly from the response stream
                        Workbook workbook = new Workbook(responseStream);

                        // List all worksheet names
                        Console.WriteLine("Worksheets in the workbook:");
                        foreach (Worksheet sheet in workbook.Worksheets)
                        {
                            Console.WriteLine("- " + sheet.Name);
                        }

                        // Read a cell value from the first worksheet
                        Worksheet firstSheet = workbook.Worksheets[0];
                        Cell cellA1 = firstSheet.Cells["A1"];
                        Console.WriteLine($"Value of A1 in '{firstSheet.Name}': {cellA1.StringValue}");
                    }
                }
            }
        }
        catch (HttpRequestException ex)
        {
            Console.WriteLine($"HTTP request error: {ex.Message}");
        }
        catch (IOException ex)
        {
            Console.WriteLine($"I/O error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }

        // No saving to disk is performed; all operations are in-memory.
    }
}
