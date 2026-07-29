// Title: C# – Load JSON from REST API and Fill Excel Template with Aspose.Cells Smart Markers
// Description: This example shows how to fetch JSON from a REST endpoint (with a fallback sample), optionally deserialize it into a C# model, load an Excel workbook that contains smart markers, assign the JSON string as a data source using WorkbookDesigner.SetJsonDataSource, process the markers, and save the populated file. Robust error handling for HTTP failures and deserialization is included.
// Keywords: Aspose.Cells | C# smart markers | WorkbookDesigner SetJsonDataSource | load JSON from web service | populate Excel from REST API | JSON deserialization C# | Excel template automation | Aspose.Cells example GitHub
// Common Searches: How to bind REST API JSON to Aspose.Cells smart markers C# | Aspose.Cells WorkbookDesigner SetJsonDataSource example | C# load JSON and fill Excel template | Smart markers JSON data source Aspose.Cells | Aspose.Cells JSON fallback sample
// Developer Intent: Retrieve JSON from a web service, optionally map it to a C# object, and merge the data into an Excel workbook using Aspose.Cells smart markers.
// Use Cases: Generate personalized employee reports by pulling employee JSON from an API and filling a pre‑designed Excel template. | Automate invoice creation by fetching order details in JSON format from a service and applying them to a smart‑marker invoice workbook. | Create city‑level demographic sheets by loading city JSON data and populating a reusable Excel template.
// AI Prompts: Write a C# method that fetches JSON from a URL, returns a fallback sample on failure, and sets it as a data source for WorkbookDesigner. | Show how to design smart markers in an Excel template that reference fields from a JSON data source named "DataSource". | Provide error‑handling code for HTTP request failures and JSON deserialization when using Aspose.Cells smart markers.

using System;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerJsonExample
{
    // Sample data model matching the expected JSON structure
    // This example shows how to fetch JSON from a REST endpoint (with a fallback sample), optionally deserialize it into a C# model, load an Excel workbook that contains smart markers, assign the JSON string as a data source using WorkbookDesigner.SetJsonDataSource, process the markers, and save the populated file. Robust error handling for HTTP failures and deserialization is included.
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
            try
            {
                // 1. Retrieve JSON from a web service (fallback to sample JSON on failure)
                string jsonUrl = "https://example.com/api/person"; // replace with actual endpoint
                string jsonString = await GetJsonAsync(jsonUrl) ?? GetSampleJson();

                // 2. (Optional) Map JSON to a C# object – demonstrates deserialization
                Person person = JsonSerializer.Deserialize<Person>(jsonString);

                // 3. Load the Excel template that contains smart markers
                //    The template should have a marker like: &=$DataSource.Name
                const string templatePath = "Template.xlsx";
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                Workbook workbook = new Workbook(templatePath);

                // 4. Initialize WorkbookDesigner with the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // 5. Set the JSON string as a data source for smart markers
                //    The name \"DataSource\" must match the marker prefix in the template
                designer.SetJsonDataSource("DataSource", jsonString);

                // 6. Process the smart markers – they will be replaced with JSON values
                designer.Process();

                // 7. Save the populated workbook
                const string resultPath = "Result.xlsx";
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to {resultPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to perform an HTTP GET and return the response body as a string
        private static async Task<string?> GetJsonAsync(string requestUri)
        {
            try
            {
                using HttpClient client = new HttpClient();
                HttpResponseMessage response = await client.GetAsync(requestUri);
                response.EnsureSuccessStatusCode();
                return await response.Content.ReadAsStringAsync();
            }
            catch (HttpRequestException httpEx)
            {
                Console.WriteLine($"HTTP request failed: {httpEx.Message}");
                return null; // Signal caller to use fallback JSON
            }
        }

        // Provides a simple JSON sample when the web request fails
        private static string GetSampleJson()
        {
            var sample = new Person
            {
                Name = "John Doe",
                Age = 30,
                City = "New York"
            };
            return JsonSerializer.Serialize(sample);
        }
    }
}
