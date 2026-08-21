// Title: Download Excel via HttpClient, apply OOXML Level 9 compression, and save with Aspose.Cells for .NET
// Description: Shows how to fetch an .xlsx file from a remote URL using HttpClient, load it into an Aspose.Cells Workbook from a network stream, configure OoxmlSaveOptions.CompressionType to Level9 for maximum OOXML compression, and write the compressed workbook to a local path.
// Keywords: Aspose.Cells | C# | HttpClient download Excel | network stream workbook | OoxmlSaveOptions | Level9 compression | OOXML compression | save compressed workbook | reduce .xlsx size | Aspose.Cells compression example
// Common Searches: Aspose.Cells set OoxmlCompressionType Level9 | C# download Excel file and compress with Aspose | How to save workbook from stream with maximum compression | Compress .xlsx using Aspose.Cells .NET | Save workbook with OOXML Level9 compression
// Developer Intent: The developer wants to retrieve an Excel file from a web URL, compress it using the highest OOXML compression level, and store the smaller file on disk.
// Use Cases: Archiving large Excel reports to minimize storage costs | Sending compressed workbooks as email attachments | Distributing lightweight Excel files over low‑bandwidth networks | Storing Excel data in cloud storage with reduced size
// AI Prompts: Generate C# code that downloads an .xlsx file with HttpClient, loads it into an Aspose.Cells Workbook, sets OoxmlSaveOptions.CompressionType to Level9, and saves the compressed file to a specified location. | Explain how OoxmlCompressionType.Level9 affects .xlsx file size and save performance in Aspose.Cells.

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsCompressionDemo
{
    // Shows how to fetch an .xlsx file from a remote URL using HttpClient, load it into an Aspose.Cells Workbook from a network stream, configure OoxmlSaveOptions.CompressionType to Level9 for maximum OOXML compression, and write the compressed workbook to a local path.
    class Program
    {
        static async Task Main(string[] args)
        {
            // URL of the Excel file to download
            string fileUrl = "https://example.com/sample.xlsx";

            // Local path where the compressed workbook will be saved
            string outputPath = "compressed_output.xlsx";

            try
            {
                using HttpClient httpClient = new HttpClient();

                // Send request and ensure a successful response
                using HttpResponseMessage response = await httpClient.GetAsync(fileUrl);
                response.EnsureSuccessStatusCode();

                // Read the content as a stream
                await using Stream networkStream = await response.Content.ReadAsStreamAsync();

                // Load the workbook from the network stream
                using Workbook workbook = new Workbook(networkStream);

                // Configure OOXML save options with maximum compression (Level9)
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    CompressionType = OoxmlCompressionType.Level9
                };

                // Ensure the output directory exists
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook with the specified compression options
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook downloaded, compressed with Level9, and saved to '{outputPath}'.");
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"Error downloading the file: {ex.Message}");
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
}
