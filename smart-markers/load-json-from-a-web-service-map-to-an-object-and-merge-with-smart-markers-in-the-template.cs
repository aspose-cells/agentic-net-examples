// Title: Download JSON from a REST endpoint and populate an Excel template with smart markers using Aspose.Cells WorkbookDesigner in C#
// AI Prompts: Fetch JSON from a specified URL using HttpClient, assign the JSON string to a WorkbookDesigner with the data source name "DataSource", process all smart markers in the .xlsx template, and save the resulting workbook. | Deserialize the retrieved JSON into a strongly‑typed C# object, then use the same JSON payload as the data source for Aspose.Cells smart markers to generate a filled Excel file.
// Common Searches: C# Aspose.Cells WorkbookDesigner SetJsonDataSource from web service | How to fill Excel smart markers with JSON returned by a REST API | Aspose.Cells example loading JSON and merging into template.xlsx | Populate Excel template using smart markers and JSON data in .NET
// Tags: Aspose.Cells JSON data source for smart markers | WorkbookDesigner smart marker population from web API | C# download and merge JSON into Excel template | Excel template processing with Aspose.Cells and JSON | Smart markers integration with REST JSON in .NET

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsJsonSmartMarkerDemo
{
    // Sample data class that matches the JSON structure (optional mapping)
    // The sample program downloads JSON from a web service, optionally deserializes it into a Person object, loads an Excel workbook that contains smart markers, sets the JSON string as a data source named "DataSource" via WorkbookDesigner, processes the markers, and saves the populated workbook to a new file.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    public class Program
    {
        // Entry point
        public static async Task Main()
        {
            try
            {
                // URL of the web service that returns JSON data
                string jsonUrl = "https://example.com/api/person";

                // Download JSON string from the web service (fallback to sample JSON on failure)
                string jsonData = await DownloadJsonAsync(jsonUrl);

                // (Optional) Map JSON to a strongly‑typed object – this step demonstrates deserialization
                // If you only need to merge with smart markers you can skip this and use jsonData directly.
                Person person = System.Text.Json.JsonSerializer.Deserialize<Person>(jsonData);

                // Path to the Excel template that contains smart markers (e.g., &=$DataSource.Name)
                string templatePath = "TemplateWithSmartMarkers.xlsx";

                // Ensure the template file exists before loading
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file \"{templatePath}\" not found.");
                    return;
                }

                // Load the Excel template
                Workbook workbook = new Workbook(templatePath);

                // Create a WorkbookDesigner and assign the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Set the JSON data source for the smart markers.
                // The name "DataSource" must match the name used in the smart markers inside the template.
                designer.SetJsonDataSource("DataSource", jsonData);

                // Process all smart markers in the workbook
                designer.Process();

                // Save the populated workbook
                string resultPath = "ResultFromWebService.xlsx";
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved to \"{resultPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to download JSON content using HttpClient
        private static async Task<string> DownloadJsonAsync(string requestUri)
        {
            try
            {
                using HttpClient client = new HttpClient();
                HttpResponseMessage response = await client.GetAsync(requestUri);
                response.EnsureSuccessStatusCode();
                return await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                // Log the error and provide a fallback JSON string
                Console.WriteLine($"Failed to download JSON from \"{requestUri}\": {ex.Message}");
                // Sample fallback JSON matching the Person class
                return @"{ ""Name"": ""John Doe"", ""Age"": 30, ""City"": ""New York"" }";
            }
        }
    }
}
