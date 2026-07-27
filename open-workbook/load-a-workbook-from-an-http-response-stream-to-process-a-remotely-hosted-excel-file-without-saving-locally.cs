// Title: Load an Excel workbook from a URL directly into Aspose.Cells (.NET) without disk I/O
// Description: Demonstrates how to download an XLSX file with HttpClient, create an Aspose.Cells Workbook from the response stream, read cell values, and save the result to a MemoryStream. Includes a fallback to a local file when the download fails and proper disposal of resources.
// Keywords: Aspose.Cells load workbook from stream | C# open Excel from URL | HttpClient Excel download Aspose | process remote XLSX without saving | Aspose.Cells memory stream save | fallback to local Excel file | in‑memory Excel manipulation .NET
// Common Searches: How to open an Excel file from a web URL using Aspose.Cells C# | Aspose.Cells read cell from HttpClient response stream | Load workbook from remote location without writing to disk | C# Aspose.Cells fallback to local file if download fails | Save Aspose.Cells workbook to MemoryStream
// Developer Intent: Open an Excel workbook directly from an HTTP response stream, manipulate it, and keep all operations in memory.
// Use Cases: Read specific cells (e.g., A1) from an Excel file delivered by a web service. | Process a remote spreadsheet in a server‑less or sandboxed environment where file‑system access is restricted. | Automatically switch to a local copy when the remote download is unavailable. | Transmit the modified workbook over a network by saving it to a MemoryStream.
// AI Prompts: Write C# code that uses Aspose.Cells to load an XLSX file from an HttpClient response stream, reads cell A1, and returns the workbook as a MemoryStream. | Show how to add a fallback that loads a local Excel file if the HTTP download fails, using Aspose.Cells. | Provide an example of correctly disposing an Aspose.Cells Workbook after processing a workbook loaded from a network stream.

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsHttpLoadDemo
{
    // Demonstrates how to download an XLSX file with HttpClient, create an Aspose.Cells Workbook from the response stream, read cell values, and save the result to a MemoryStream. Includes a fallback to a local file when the download fails and proper disposal of resources.
    class Program
    {
        static async Task Main(string[] args)
        {
            string excelUrl = "https://example.com/sample.xlsx";
            Workbook workbook = null;

            try
            {
                // Attempt to download the workbook from the remote URL
                using var httpClient = new HttpClient();
                using var response = await httpClient.GetAsync(excelUrl, HttpCompletionOption.ResponseHeadersRead);
                response.EnsureSuccessStatusCode();

                using var excelStream = await response.Content.ReadAsStreamAsync();
                workbook = new Workbook(excelStream);
                Console.WriteLine("Workbook loaded from remote URL.");
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"Failed to download workbook: {ex.Message}");

                // Fallback to a local file if it exists
                string localPath = "sample.xlsx";
                if (File.Exists(localPath))
                {
                    try
                    {
                        workbook = new Workbook(localPath);
                        Console.WriteLine("Workbook loaded from local file.");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"Failed to load local workbook: {e.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Local file '{localPath}' not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }

            if (workbook != null)
            {
                try
                {
                    // Example processing: read the value of cell A1 from the first worksheet
                    Worksheet firstSheet = workbook.Worksheets[0];
                    string cellValue = firstSheet.Cells["A1"].StringValue;
                    Console.WriteLine($"Value of A1 in the workbook: {cellValue}");

                    // Save the workbook to a memory stream (no file system access)
                    using var memoryStream = new MemoryStream();
                    workbook.Save(memoryStream, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to memory stream. Length = {memoryStream.Length} bytes.");
                }
                finally
                {
                    // Ensure resources are released
                    workbook.Dispose();
                }
            }
        }
    }
}
