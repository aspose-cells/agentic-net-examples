// Title: Load an Excel workbook from an HTTP response stream with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to use HttpClient to download an Excel file, pass the response stream directly to the Aspose.Cells Workbook constructor, read cell values, and optionally save the workbook to a MemoryStream for in‑memory processing without writing to disk.
// Keywords: Aspose.Cells | load workbook from stream | C# HttpClient Excel | read remote Excel without file | Workbook constructor stream | in‑memory Excel processing
// Common Searches: Aspose.Cells open Excel from URL C# | load workbook from HttpResponseMessage stream | read cell A1 from remote Excel using Aspose | process Excel file in memory .NET | download Excel with HttpClient and Aspose.Cells
// Developer Intent: Load a remote Excel workbook directly from an HTTP response stream for immediate processing without persisting the file.
// Use Cases: Extract specific cell values from a spreadsheet hosted on an external server. | Convert a downloaded workbook to a MemoryStream for API responses or further transformations. | Integrate on‑the‑fly Excel validation or conversion into a web service that consumes URLs.
// AI Prompts: Generate C# code that uses Aspose.Cells to open an Excel file from an HttpResponseMessage stream and iterates over all rows in the first worksheet. | Show how to download an Excel file with HttpClient, load it into a Workbook, modify a cell, and return the updated file as a byte array without writing to disk. | Explain best‑practice error handling when loading a workbook from a network stream with Aspose.Cells, including timeout and retry strategies.

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// Demonstrates how to use HttpClient to download an Excel file, pass the response stream directly to the Aspose.Cells Workbook constructor, read cell values, and optionally save the workbook to a MemoryStream for in‑memory processing without writing to disk.
class LoadWorkbookFromHttpStream
{
    static async Task Main(string[] args)
    {
        // URL of the remote Excel file (replace with a valid URL)
        string excelUrl = "https://example.com/sample.xlsx";

        try
        {
            // Initialize HttpClient (should be reused in real applications)
            using var httpClient = new HttpClient();

            // Send GET request and obtain the response stream without saving to disk
            using HttpResponseMessage response = await httpClient.GetAsync(
                excelUrl, HttpCompletionOption.ResponseHeadersRead);

            if (!response.IsSuccessStatusCode)
            {
                Console.WriteLine($"Failed to download file. Status code: {(int)response.StatusCode} {response.ReasonPhrase}");
                return;
            }

            // Read the content as a stream
            await using Stream httpStream = await response.Content.ReadAsStreamAsync();

            // Load the workbook directly from the HTTP response stream
            var workbook = new Workbook(httpStream);

            // Example processing: read the value of cell A1 from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].StringValue}");

            // Optionally, save the workbook to a memory stream (e.g., for further transmission)
            await using var memoryStream = new MemoryStream();
            workbook.Save(memoryStream, SaveFormat.Xlsx);
            memoryStream.Position = 0; // Reset for any subsequent reading

            // Example: write the size of the in‑memory workbook
            Console.WriteLine($"Workbook saved to memory stream, length = {memoryStream.Length} bytes");
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
    }
}
