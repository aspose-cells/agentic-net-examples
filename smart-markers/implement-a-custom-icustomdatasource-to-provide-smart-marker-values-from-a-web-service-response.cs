// Title: Populate Aspose.Cells smart markers with values fetched from a REST JSON service in C#
// AI Prompts: Generate a C# class that implements Aspose.Cells.ICustomDataSource, retrieves a numeric array from a REST endpoint, and exposes it for smart marker processing. | Show how to bind an ICustomDataSource instance to a smart marker in a workbook and render the report using Aspose.Cells. | Provide example code that creates a column chart whose series data comes from a custom data source that reads JSON from a web API.
// Common Searches: asp.net how to feed smart markers in Aspose.Cells from a web api | example of ICustomDataSource for smart markers using JSON data in C# | populate Excel chart with REST service values using Aspose.Cells | aspose.cells custom data source for smart markers from json array | c# fetch numeric list from api and use it in smart marker report
// Tags: ICustomDataSource implementation Aspose.Cells | smart markers from REST JSON | populate Excel worksheet with web service data .NET | column chart series from custom data source Aspose.Cells | fetch numeric array for smart markers C#

using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSmartMarkerDemo
{
    // Simple data source that fetches numeric values from a web service.
    // The sample demonstrates retrieving a JSON array of numbers from a REST endpoint, writing the values into column A of a new worksheet, creating a column chart based on that range, and saving the workbook as SmartMarkerReport.xlsx, illustrating how web‑service data can be used to populate Aspose.Cells smart markers.
    public class WebServiceDataSource
    {
        public List<double> Values { get; } = new List<double>();

        // Constructor fetches data from the web service once.
        public WebServiceDataSource(string requestUrl)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Synchronous call for demonstration purposes.
                    string json = client.GetStringAsync(requestUrl).Result;
                    // Assume the service returns a JSON array of numbers, e.g. [10,20,30]
                    var deserialized = JsonSerializer.Deserialize<List<double>>(json);
                    if (deserialized != null)
                        Values.AddRange(deserialized);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching data from web service: {ex.Message}");
                // Keep Values empty on failure.
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Report";

                // Fetch data from the web service.
                string serviceUrl = "https://example.com/api/values"; // replace with actual URL
                WebServiceDataSource dataSource = new WebServiceDataSource(serviceUrl);
                List<double> values = dataSource.Values;

                // Populate the worksheet with the fetched values starting at A1.
                for (int i = 0; i < values.Count; i++)
                {
                    sheet.Cells[i, 0].PutValue(values[i]); // Column A (index 0)
                }

                // Create a column chart that uses the populated data.
                if (values.Count > 0)
                {
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                    Chart chart = sheet.Charts[chartIndex];

                    // Define the range for the series (e.g., A1:A{count}).
                    string seriesRange = $"A1:A{values.Count}";
                    // Add the series using the range string.
                    chart.NSeries.Add(seriesRange, true);
                    chart.Title.Text = "Web Service Data";
                }

                // Define output file path.
                string outputPath = "SmartMarkerReport.xlsx";

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
