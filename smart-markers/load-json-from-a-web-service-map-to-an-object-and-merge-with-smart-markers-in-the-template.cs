// Title: C# – Load JSON from a Web Service and Populate Excel Smart Markers with Aspose.Cells
// Description: Fetch JSON via HttpClient, optionally deserialize to a C# model, set it as a JSON data source for WorkbookDesigner, process smart markers in an Excel template (auto‑created if missing), and save the populated workbook.
// Keywords: Aspose.Cells | WorkbookDesigner | SetJsonDataSource | smart markers | C# JSON web service | HttpClient | JSON deserialization | Excel template | fallback JSON | .NET
// Common Searches: Aspose.Cells set JSON data source example | C# smart markers from web service | WorkbookDesigner populate Excel from JSON | Create Excel template with smart markers programmatically | Handle HTTP errors with fallback JSON Aspose.Cells
// Developer Intent: Retrieve JSON from a URL, map it to a .NET object, and merge the data into Excel smart markers using Aspose.Cells.
// Use Cases: Generate employee or customer reports by pulling data from a REST API and filling an Excel template with smart markers. | Automatically create a minimal Excel template with smart markers when the expected file is missing, then populate it with live JSON data. | Provide a resilient workflow that falls back to a hard‑coded JSON payload if the web request fails, ensuring the workbook is still produced.
// AI Prompts: Write C# code that uses Aspose.Cells WorkbookDesigner to set a JSON data source from an HttpClient response and process smart markers. | Show how to programmatically create a simple Excel template containing smart markers when the template file does not exist. | Explain how to deserialize a JSON response into a strongly‑typed C# class and then use SetJsonDataSource to fill smart markers.

using System;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsJsonSmartMarkerDemo
{
    // Sample data model matching the JSON structure
    // Fetch JSON via HttpClient, optionally deserialize to a C# model, set it as a JSON data source for WorkbookDesigner, process smart markers in an Excel template (auto‑created if missing), and save the populated workbook.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    class Program
    {
        // Entry point – async to allow awaiting the HTTP call
        static async Task Main(string[] args)
        {
            // URL of the web service returning JSON data
            const string jsonUrl = "https://example.com/api/person";

            string jsonData = string.Empty;

            // Retrieve JSON string from the web service with error handling
            try
            {
                using HttpClient client = new HttpClient();
                jsonData = await client.GetStringAsync(jsonUrl);
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"Warning: Unable to retrieve JSON from '{jsonUrl}'. {ex.Message}");
                // Fallback to a sample JSON payload
                var samplePerson = new Person { Name = "John Doe", Age = 30, City = "New York" };
                jsonData = JsonSerializer.Serialize(samplePerson);
                Console.WriteLine("Using fallback JSON data.");
            }

            // Optional: map JSON to a strongly‑typed object (demonstration purpose)
            try
            {
                Person person = JsonSerializer.Deserialize<Person>(jsonData);
                Console.WriteLine($"Deserialized Person: {person?.Name}, {person?.Age}, {person?.City}");
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"Error deserializing JSON: {ex.Message}");
            }

            const string templatePath = "Template.xlsx";

            // Ensure the template file exists; if not, create a minimal workbook with a smart marker
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file '{templatePath}' not found. Creating a default template.");
                Workbook tempWb = new Workbook();
                Worksheet sheet = tempWb.Worksheets[0];
                // Insert a smart marker that matches the data source name ("DataSource")
                sheet.Cells["A1"].PutValue("&=$DataSource.Name");
                sheet.Cells["A2"].PutValue("&=$DataSource.Age");
                sheet.Cells["A3"].PutValue("&=$DataSource.City");
                tempWb.Save(templatePath);
            }

            // Load the Excel template that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set the JSON string as a data source for smart markers.
            // The name "DataSource" must match the marker prefix used in the template.
            designer.SetJsonDataSource("DataSource", jsonData);

            // Process all smart markers in the workbook
            try
            {
                designer.Process();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing smart markers: {ex.Message}");
                return;
            }

            // Save the populated workbook
            const string resultPath = "Result.xlsx";
            try
            {
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
