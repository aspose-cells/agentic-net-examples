// Title: Populate Aspose.Cells smart markers from a REST API using a DataTable in C#
// Common Searches: C# Aspose.Cells set smart marker data source from REST API JSON | How to use WorkbookDesigner SetDataSource with a DataTable retrieved via HTTP GET | Example of filling Excel smart markers with web service data in .NET | Load JSON array into a DataTable and bind it to Aspose.Cells smart markers
// Tags: Aspose.Cells WorkbookDesigner SetDataSource DataTable | C# convert JSON array to DataTable for Excel | populate smart markers from REST API data | Excel template automation with web service JSON | Aspose.Cells smart markers using HTTP GET response

using System;
using System.Data;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsSmartMarkersExample
{
    // The program loads a template workbook, fetches JSON data from a REST endpoint, converts the JSON array into a DataTable, assigns this table as the "RestApiData" source for Aspose.Cells smart markers via WorkbookDesigner, processes the markers, and saves the populated workbook.
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string outputPath = "Output.xlsx";
                const string apiUrl = "https://api.example.com/users";

                // Verify that the template file exists
                if (!File.Exists(templatePath))
                {
                    Console.Error.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook
                var workbook = new Workbook(templatePath);

                // Retrieve data from the REST API as a DataTable
                DataTable apiData = await FetchDataTableFromApiAsync(apiUrl);

                // Set the data source for smart markers
                var designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("RestApiData", apiData);
                designer.Process();

                // Save the populated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Retrieves JSON from the API and converts it to a DataTable
        private static async Task<DataTable> FetchDataTableFromApiAsync(string apiUrl)
        {
            using var client = new HttpClient();
            HttpResponseMessage response = await client.GetAsync(apiUrl);
            response.EnsureSuccessStatusCode();

            string json = await response.Content.ReadAsStringAsync();

            // Expecting a JSON array of objects, e.g., [{ "Name":"John","Age":30 }, ...]
            using JsonDocument doc = JsonDocument.Parse(json);
            JsonElement root = doc.RootElement;

            if (root.ValueKind != JsonValueKind.Array)
                throw new InvalidOperationException("Expected JSON array from API.");

            var table = new DataTable();
            bool columnsDefined = false;

            foreach (JsonElement element in root.EnumerateArray())
            {
                if (!columnsDefined)
                {
                    // Define columns based on the first object's properties
                    foreach (JsonProperty prop in element.EnumerateObject())
                    {
                        table.Columns.Add(prop.Name, typeof(string));
                    }
                    columnsDefined = true;
                }

                var row = table.NewRow();
                foreach (JsonProperty prop in element.EnumerateObject())
                {
                    // Store the raw JSON value as string (trim surrounding quotes if any)
                    row[prop.Name] = prop.Value.GetRawText().Trim('\"');
                }
                table.Rows.Add(row);
            }

            return table;
        }
    }
}
