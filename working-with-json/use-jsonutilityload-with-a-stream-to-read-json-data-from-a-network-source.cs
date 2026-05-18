using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonFromNetwork
{
    class Program
    {
        // Entry point of the console application
        static async Task Main(string[] args)
        {
            try
            {
                // URL of the JSON file hosted on a network location
                string jsonUrl = "https://example.com/data/sample.json";

                // Create HttpClient for downloading the JSON content as a stream
                using (HttpClient httpClient = new HttpClient())
                {
                    // Send GET request
                    using (HttpResponseMessage response = await httpClient.GetAsync(jsonUrl))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            Console.WriteLine($"Failed to download JSON. Status code: {(int)response.StatusCode} {response.ReasonPhrase}");
                            return;
                        }

                        // Read the response content as a stream
                        using (Stream jsonStream = await response.Content.ReadAsStreamAsync())
                        {
                            // Configure JSON load options (optional)
                            JsonLoadOptions loadOptions = new JsonLoadOptions
                            {
                                // Example: start importing data from cell B2
                                StartCell = "B2",
                                // Keep the JSON schema for later saving back to JSON if needed
                                KeptSchema = true
                            };

                            // Load the JSON data into a Workbook using the stream and options
                            Workbook workbook = new Workbook(jsonStream, loadOptions);

                            // Save the workbook to an Excel file
                            string outputPath = "NetworkJsonOutput.xlsx";
                            workbook.Save(outputPath);
                            Console.WriteLine($"JSON data loaded from network and saved to Excel successfully: {Path.GetFullPath(outputPath)}");
                        }
                    }
                }
            }
            catch (HttpRequestException httpEx)
            {
                Console.WriteLine($"Network error while downloading JSON: {httpEx.Message}");
            }
            catch (FileNotFoundException fileEx)
            {
                Console.WriteLine($"File not found: {fileEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}