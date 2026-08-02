// Title: Custom ICellsDataTable for JSON Web‑Service Smart Markers in Aspose.Cells (.NET)
// Description: Implements ICellsDataTable to parse a JSON array from a web service, expose rows and columns, and bind the source to WorkbookDesigner so smart markers like &=$WebData.Name populate an Excel workbook automatically.
// Keywords: Aspose.Cells | ICellsDataTable | smart markers | JSON web service | C# | WorkbookDesigner | custom data source | Excel report generation | .NET
// Common Searches: bind json response to aspose.cells smart markers | implement icellsdatatable c# | populate excel from web api using aspose.cells | aspose.cells custom datasource example | smart markers json array
// Developer Intent: Create a reusable ICellsDataTable that converts a JSON API response into a data table for Aspose.Cells smart markers.
// Use Cases: Generate sales or inventory reports directly from a REST endpoint without intermediate POCO models. | Create master‑detail Excel sheets by registering multiple WebServiceDataSource instances (e.g., Orders, Customers) with WorkbookDesigner. | Apply conditional formatting or grouping while iterating over the custom data source during smart‑marker processing.
// AI Prompts: Write a C# method that calls a REST API with HttpClient, builds a WebServiceDataSource from the JSON, and processes a workbook template containing smart markers. | Explain how to extend WebServiceDataSource to handle nested JSON objects and expose them as separate smart‑marker tables. | Provide sample code that registers two custom data sources ("Orders" and "Customers") with WorkbookDesigner and uses smart markers to produce a master‑detail Excel report.

using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Tables; // ICellsDataTable

// Custom data source that retrieves data from a web service (JSON response)
// Implements ICellsDataTable to parse a JSON array from a web service, expose rows and columns, and bind the source to WorkbookDesigner so smart markers like &=$WebData.Name populate an Excel workbook automatically.
public class WebServiceDataSource : ICellsDataTable
{
    // Internal list of records parsed from JSON
    private readonly List<Dictionary<string, object>> _records;
    private int _currentRow = -1;

    // Constructor accepts the JSON string (could be fetched beforehand)
    public WebServiceDataSource(string json)
    {
        // Parse JSON array of objects into a list of dictionaries
        var temp = JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(json,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new List<Dictionary<string, JsonElement>>();

        _records = temp.ConvertAll(dict =>
        {
            var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in dict)
            {
                // Convert JsonElement to appropriate .NET type
                object value = kvp.Value.ValueKind switch
                {
                    JsonValueKind.String => kvp.Value.GetString(),
                    JsonValueKind.Number => kvp.Value.TryGetInt64(out long l) ? (object)l :
                                            kvp.Value.TryGetDouble(out double d) ? (object)d : kvp.Value.GetDecimal(),
                    JsonValueKind.True => true,
                    JsonValueKind.False => false,
                    JsonValueKind.Null => null,
                    _ => kvp.Value.GetRawText()
                };
                result[kvp.Key] = value;
            }
            return result;
        });
    }

    // Indexer by row and column (zero‑based). Column order follows the Columns array.
    public object this[int rowIndex, int columnIndex]
    {
        get
        {
            string colName = Columns[columnIndex];
            return _records[rowIndex][colName];
        }
    }

    // Indexer by row only – returns the whole record as a dictionary
    public object this[int rowIndex] => _records[rowIndex];

    // Indexer by column name – works on the current row (set by Next())
    public object this[string columnName] => _records[_currentRow][columnName];

    // Total number of rows
    public int RowCount => _records.Count;

    // Total number of columns (derived from first record or empty)
    public int ColumnCount => Columns.Length;

    // Alias for RowCount (required by ICellsDataTable)
    public int Count => RowCount;

    // Column names – derived from the first record (order of appearance)
    public string[] Columns
    {
        get
        {
            if (_records.Count == 0) return Array.Empty<string>();
            var keys = new List<string>(_records[0].Keys);
            return keys.ToArray();
        }
    }

    // Reset enumeration to before the first row
    public void BeforeFirst()
    {
        _currentRow = -1;
    }

    // Move to the next row; returns false when no more rows
    public bool Next()
    {
        _currentRow++;
        return _currentRow < _records.Count;
    }
}

// Example usage demonstrating the custom data source with smart markers
public class SmartMarkerWebServiceDemo
{
    public static void Run()
    {
        try
        {
            // 1. Create a new workbook (template) and place smart markers
            Workbook workbook = new Workbook(); // create empty workbook
            Worksheet sheet = workbook.Worksheets[0];
            // Smart markers reference the fields returned by the web service
            sheet.Cells["A1"].PutValue("&=$WebData.Id");
            sheet.Cells["B1"].PutValue("&=$WebData.Name");
            sheet.Cells["C1"].PutValue("&=$WebData.Price");

            // 2. Simulate a web service call (replace with real HttpClient call if needed)
            string jsonResponse = FetchDataFromWebService().Result; // async result for brevity

            // 3. Create the custom data source using the JSON response
            var dataSource = new WebServiceDataSource(jsonResponse);

            // 4. Set up the WorkbookDesigner and bind the custom data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook // assign workbook
            };
            designer.SetDataSource("WebData", dataSource); // bind custom source

            // 5. Process smart markers and generate the final workbook
            designer.Process();

            // 6. Save the result (output)
            string outputPath = "WebServiceSmartMarkersOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error during SmartMarkerWebServiceDemo: {ex.Message}");
        }
    }

    // Mock method that pretends to call a remote API and returns JSON.
    // Replace the implementation with a real HttpClient request as needed.
    private static async Task<string> FetchDataFromWebService()
    {
        // Example JSON array returned by the service
        // [{"Id":1,"Name":"Product A","Price":12.5},{"Id":2,"Name":"Product B","Price":23.0}]
        await Task.Delay(10); // simulate async latency
        return "[{\"Id\":1,\"Name\":\"Product A\",\"Price\":12.5},{\"Id\":2,\"Name\":\"Product B\",\"Price\":23.0}]";
    }
}

// Entry point
class Program
{
    static void Main()
    {
        SmartMarkerWebServiceDemo.Run();
        Console.WriteLine("Workbook generated successfully.");
    }
}
