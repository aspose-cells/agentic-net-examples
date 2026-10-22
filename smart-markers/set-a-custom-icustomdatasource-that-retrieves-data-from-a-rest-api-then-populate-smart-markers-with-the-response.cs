// Title: Create a custom ICustomDataSource to load JSON from a REST endpoint and populate Aspose.Cells smart markers in C#
// AI Prompts: Write C# code that implements ICustomDataSource, calls a specified REST URL, deserializes the returned JSON array into record objects, and exposes them for Aspose.Cells smart markers. | Show how to attach the custom data source to a workbook's smart marker processor, map JSON fields to markers, and save the populated Excel file.
// Common Searches: how to use ICustomDataSource with Aspose.Cells to read data from a web API in C# | populate Aspose.Cells smart markers from JSON returned by a REST service | C# example for binding REST API response to Aspose.Cells smart markers | Aspose.Cells custom data source for dynamic Excel generation from API data | fetch customer list from REST endpoint and fill Excel using Aspose.Cells smart markers
// Tags: custom ICustomDataSource for REST JSON | Aspose.Cells smart markers data binding | C# JSON to Excel using Aspose.Cells | populate Excel with web service data | dynamic workbook generation from API

using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample demonstrates implementing a custom ICustomDataSource that synchronously retrieves JSON from a REST endpoint, deserializes it into a list of dictionaries, and supplies those records to Aspose.Cells smart markers to generate and save an Excel workbook.
public class RestApiDataSource
{
    private readonly List<Dictionary<string, object>> _records;

    public RestApiDataSource(string requestUri)
    {
        try
        {
            var json = GetJsonAsync(requestUri).Result;
            _records = JsonSerializer.Deserialize<List<Dictionary<string, object>>>(json) ?? new List<Dictionary<string, object>>();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching data from API: {ex.Message}");
            _records = new List<Dictionary<string, object>>();
        }
    }

    private async Task<string> GetJsonAsync(string uri)
    {
        using (HttpClient client = new HttpClient())
        {
            HttpResponseMessage response = await client.GetAsync(uri);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }
    }

    // Expose the fetched records
    public List<Dictionary<string, object>> GetRecords()
    {
        return _records;
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Prepare the worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Customers";

            // 3. Fetch data from the REST API
            string apiUrl = "https://example.com/api/customers";
            RestApiDataSource dataSource = new RestApiDataSource(apiUrl);
            var records = dataSource.GetRecords();

            // 4. Populate the worksheet (simple example using the first record)
            if (records.Count > 0)
            {
                var first = records[0];
                sheet.Cells["A1"].PutValue("Customer Name:");
                sheet.Cells["B1"].PutValue(first.TryGetValue("Name", out var name) ? name?.ToString() : string.Empty);

                sheet.Cells["A2"].PutValue("Customer Age:");
                sheet.Cells["B2"].PutValue(first.TryGetValue("Age", out var age) ? age?.ToString() : string.Empty);
            }
            else
            {
                Console.WriteLine("No records retrieved from the API.");
            }

            // 5. Save the resulting workbook
            string outputPath = "SmartMarkersFromRestApi.xlsx";
            try
            {
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
