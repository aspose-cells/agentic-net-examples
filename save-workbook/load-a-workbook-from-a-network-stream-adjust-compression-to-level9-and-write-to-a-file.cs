using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

namespace AsposeCellsNetworkExample
{
    class Program
    {
        static void Main()
        {
            // URL of the Excel file to load from the network
            const string excelUrl = "https://example.com/sample.xlsx";

            // Destination file path
            const string outputPath = "DownloadedWithCompression.xlsx";

            try
            {
                // Download the file into a memory stream
                using (HttpClient httpClient = new HttpClient())
                {
                    // Get the response and ensure success status
                    HttpResponseMessage response = httpClient.GetAsync(excelUrl).Result;
                    response.EnsureSuccessStatusCode();

                    using (Stream networkStream = response.Content.ReadAsStreamAsync().Result)
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        // Copy the network stream to a seekable memory stream
                        networkStream.CopyTo(memoryStream);
                        memoryStream.Position = 0; // Reset position for reading

                        // Load the workbook from the memory stream
                        Workbook workbook = new Workbook(memoryStream);

                        // Create OOXML save options and set maximum compression (Level9)
                        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                        {
                            CompressionType = OoxmlCompressionType.Level9
                        };

                        // Save the workbook to a file using the specified compression options
                        workbook.Save(outputPath, saveOptions);
                    }
                }

                Console.WriteLine("Workbook downloaded, compressed with Level9, and saved to file.");
            }
            catch (HttpRequestException httpEx)
            {
                Console.WriteLine($"Network error while downloading the file: {httpEx.Message}");
            }
            catch (IOException ioEx)
            {
                Console.WriteLine($"File I/O error: {ioEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}