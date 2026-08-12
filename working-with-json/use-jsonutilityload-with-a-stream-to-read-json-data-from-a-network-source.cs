// Title: Load JSON from a Web URL into an Aspose.Cells Workbook via Stream with JsonLoadOptions (C#)
// Description: C# example that downloads a JSON file using HttpClient, reads it as a stream, and imports the data directly into an Aspose.Cells Workbook with JsonLoadOptions (StartCell B2, KeptSchema enabled, ArrayAsTable layout). The workbook is verified and saved as an XLSX file.
// Keywords: Aspose.Cells JSON import | JsonUtility.Load stream | C# HttpClient download JSON | JsonLoadOptions StartCell | ArrayAsTable layout | load JSON from URL | Excel export from JSON | Aspose.Cells network JSON
// Common Searches: How to import JSON from a web service into Aspose.Cells | Aspose.Cells JsonUtility.Load with HttpClient example | Set start cell when loading JSON in Aspose.Cells | Treat JSON arrays as tables in Aspose.Cells | Load JSON stream into workbook C#
// Developer Intent: Download a remote JSON file and load it directly into an Aspose.Cells workbook using a stream and configurable JsonLoadOptions.
// Use Cases: Convert API response JSON to Excel for quick reporting. | Import server‑hosted configuration data into a spreadsheet for analysis. | Preserve JSON schema while transforming data to XLSX for downstream processing.
// AI Prompts: Show how to process very large JSON files with buffered streaming in Aspose.Cells. | Provide a version that uses JsonUtility.Load instead of the Workbook constructor while keeping the same load options. | Explain how to map nested JSON objects to separate worksheets using JsonLayoutOptions.

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonFromNetwork
{
    // C# example that downloads a JSON file using HttpClient, reads it as a stream, and imports the data directly into an Aspose.Cells Workbook with JsonLoadOptions (StartCell B2, KeptSchema enabled, ArrayAsTable layout). The workbook is verified and saved as an XLSX file.
    class Program
    {
        static async Task Main(string[] args)
        {
            // URL of the JSON file hosted on a network location
            string jsonUrl = "https://example.com/data/sample.json";

            try
            {
                // Initialize HttpClient for downloading the JSON content
                using (HttpClient httpClient = new HttpClient())
                {
                    // Send request and verify response status
                    HttpResponseMessage response = await httpClient.GetAsync(jsonUrl);
                    if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine($"Failed to download JSON. Status code: {response.StatusCode}");
                        return;
                    }

                    // Download the JSON data as a stream
                    using (Stream jsonStream = await response.Content.ReadAsStreamAsync())
                    {
                        // Configure JSON load options (optional customizations)
                        JsonLoadOptions loadOptions = new JsonLoadOptions
                        {
                            // Example: start importing data at cell B2
                            StartCell = "B2",
                            // Keep the JSON schema for later export if needed
                            KeptSchema = true,
                            // Use layout options to treat arrays as tables
                            LayoutOptions = new JsonLayoutOptions
                            {
                                ArrayAsTable = true
                            }
                        };

                        // Load the JSON data into a workbook using the stream and options
                        Workbook workbook = new Workbook(jsonStream, loadOptions);

                        // (Optional) Access some data to verify loading
                        Worksheet sheet = workbook.Worksheets[0];
                        Console.WriteLine("First cell value after load: " + sheet.Cells["B2"].StringValue);

                        // Save the workbook to an Excel file
                        string outputPath = "NetworkJsonOutput.xlsx";
                        workbook.Save(outputPath, SaveFormat.Xlsx);
                        Console.WriteLine($"Workbook saved to {outputPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
