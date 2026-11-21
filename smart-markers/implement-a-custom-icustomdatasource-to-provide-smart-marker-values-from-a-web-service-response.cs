// Title: Create a custom ICustomDataSource in C# that loads JSON from a web service for Aspose.Cells smart markers and binds it to a chart
// Common Searches: C# custom ICustomDataSource for Aspose.Cells smart markers pulling data from a REST API | populate Aspose.Cells worksheet with JSON array returned by a web service | bind dynamic data range to Aspose.Cells chart series using smart markers | fetch JSON data synchronously in Aspose.Cells example for chart creation | use HttpClient to retrieve data for Aspose.Cells smart marker chart in .NET
// Tags: custom ICustomDataSource for Aspose.Cells smart markers | retrieve JSON data with HttpClient for Excel chart | populate worksheet cells from web service response C# | bind chart series to dynamic range Aspose.Cells | smart markers chart from REST API data

using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example defines a SmartMarkerDataSource class that synchronously fetches a JSON array of arrays from a given URL using HttpClient, parses the values into a List<List<object>>, and exposes them via GetAllData. In Main, a Workbook and Worksheet are created, a column chart is added, the fetched data is written to cells starting at A1, the resulting range is bound to the chart series, a title is set, and the workbook is saved as an XLSX file.
public class SmartMarkerDataSource
{
    // Internal storage for the fetched data (rows x columns).
    private readonly List<List<object>> _data = new List<List<object>>();

    // Constructor initiates the async fetch and waits for completion.
    public SmartMarkerDataSource(string requestUrl)
    {
        // Synchronously wait for the async operation (for simplicity in this example).
        Task.Run(() => FetchDataAsync(requestUrl)).GetAwaiter().GetResult();
    }

    // Asynchronously calls the web service, parses JSON array of arrays, and fills _data.
    private async Task FetchDataAsync(string url)
    {
        try
        {
            using (HttpClient client = new HttpClient())
            {
                // Example: the service returns JSON like [[1,2,3],[4,5,6],...]
                string json = await client.GetStringAsync(url);

                // Simple JSON parsing without external libraries (assumes well‑formed numeric data).
                // For production code use Json.NET or System.Text.Json.
                json = json.Trim('[', ']');
                string[] rows = json.Split(new[] { "]," }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string row in rows)
                {
                    string cleanedRow = row.Trim('[', ']');
                    string[] cells = cleanedRow.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    List<object> rowData = new List<object>();
                    foreach (string cell in cells)
                    {
                        // Try to parse as double; if fails keep as string.
                        if (double.TryParse(cell, out double num))
                            rowData.Add(num);
                        else
                            rowData.Add(cell.Trim('\"'));
                    }
                    _data.Add(rowData);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching data from '{url}': {ex.Message}");
        }
    }

    // Exposes the fetched data.
    public List<List<object>> GetAllData()
    {
        return _data;
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // 1. Create a new workbook.
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Add a chart of type Column.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // 4. Retrieve data from the web service.
            string serviceUrl = "https://example.com/api/chartdata";
            SmartMarkerDataSource dataSource = new SmartMarkerDataSource(serviceUrl);
            List<List<object>> data = dataSource.GetAllData();

            if (data.Count == 0)
            {
                Console.WriteLine("No data retrieved from the web service.");
                return;
            }

            // 5. Write the data to the worksheet starting at cell A1.
            for (int i = 0; i < data.Count; i++)
            {
                List<object> row = data[i];
                for (int j = 0; j < row.Count; j++)
                {
                    sheet.Cells[i, j].PutValue(row[j]);
                }
            }

            // 6. Determine the range of the data (e.g., A1:C3).
            int rows = data.Count;
            int cols = data[0].Count;
            string startCell = CellsHelper.CellIndexToName(0, 0);
            string endCell = CellsHelper.CellIndexToName(rows - 1, cols - 1);
            string dataRange = $"{startCell}:{endCell}";

            // 7. Bind the data range to the chart series (by rows).
            chart.NSeries.Add(dataRange, true);

            // 8. Optionally set chart title.
            chart.Title.Text = "Smart Marker Chart from Web Service";

            // 9. Save the workbook.
            string outputPath = "SmartMarkerChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
