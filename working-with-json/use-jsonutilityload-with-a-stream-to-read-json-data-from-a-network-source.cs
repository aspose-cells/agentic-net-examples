using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsJsonNetworkExample
{
    class Program
    {
        // Entry point of the console application
        static async Task Main(string[] args)
        {
            // URL of the JSON file hosted on a network location
            string jsonUrl = "https://example.com/data/sample.json";

            try
            {
                // Create an HttpClient instance for downloading the JSON data
                using (HttpClient httpClient = new HttpClient())
                {
                    // Send request and get response
                    HttpResponseMessage response = await httpClient.GetAsync(jsonUrl);
                    if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine($"Failed to download JSON. Status code: {(int)response.StatusCode} {response.ReasonPhrase}");
                        return;
                    }

                    // Download the JSON content as a stream
                    using (Stream jsonStream = await response.Content.ReadAsStreamAsync())
                    {
                        // Configure JSON load options (optional)
                        JsonLoadOptions loadOptions = new JsonLoadOptions
                        {
                            // Start loading data from cell A1
                            StartCell = "A1",
                            // Keep the JSON schema (useful if you plan to save back to JSON)
                            KeptSchema = true,
                            // Import each top‑level array as a separate worksheet
                            MultipleWorksheets = true
                        };

                        // Load the JSON data from the stream into a Workbook
                        Workbook workbook = new Workbook(jsonStream, loadOptions);

                        // For demonstration, write the name of the first worksheet to the console.
                        Console.WriteLine("Workbook loaded. First worksheet name: " + workbook.Worksheets[0].Name);

                        // Save the workbook to an Excel file
                        string outputPath = "NetworkJsonOutput.xlsx";
                        workbook.Save(outputPath, SaveFormat.Xlsx);
                        Console.WriteLine($"Workbook saved to '{outputPath}'.");
                    }
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors (e.g., network issues, Aspose.Cells errors)
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}