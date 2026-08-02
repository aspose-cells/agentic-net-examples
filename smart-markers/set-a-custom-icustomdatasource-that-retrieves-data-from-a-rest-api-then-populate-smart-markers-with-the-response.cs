// Title: C# – Use a custom ICellsDataTable to feed REST‑API JSON into Aspose.Cells smart markers
// Description: Demonstrates how to implement ICellsDataTable that pulls a JSON array from a REST endpoint, converts each object into rows and columns, binds the source to WorkbookDesigner, inserts smart markers (e.g., &=$ApiData.Name) and generates a populated Excel file.
// Keywords: Aspose.Cells custom data source | ICellsDataTable REST API | smart markers JSON | C# Excel template from web service | WorkbookDesigner SetDataSource example | populate Excel with API data | Aspose.Cells REST integration
// Common Searches: Aspose.Cells bind REST API to smart markers | Implement ICellsDataTable for JSON in .NET | How to use WorkbookDesigner with custom data source | Generate Excel from web service using Aspose.Cells | C# smart markers from API response
// Developer Intent: Retrieve data from a web service and feed it directly to Aspose.Cells smart‑marker processing via a custom ICellsDataTable implementation.
// Use Cases: Create an employee directory Excel sheet by calling a people‑API and mapping Name, Age, City fields to smart markers. | Build a daily sales report where sales figures are fetched from a REST service and inserted into a pre‑designed template. | Export inventory listings from an online catalog into a formatted workbook using smart markers for SKU, Quantity, and Price.
// AI Prompts: Write a unit test for RestApiDataSource that validates handling of an empty JSON array and proper column detection. | Refactor RestApiDataSource to use async/await with HttpClient while keeping the ICellsDataTable contract intact. | Extend RestApiDataSource to support paginated APIs, aggregating multiple pages into a single data table for smart‑marker processing.

using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsRestApiDemo
{
    // Custom data source that fetches JSON from a REST API and presents it as a table for smart markers
    // Demonstrates how to implement ICellsDataTable that pulls a JSON array from a REST endpoint, converts each object into rows and columns, binds the source to WorkbookDesigner, inserts smart markers (e.g., &=$ApiData.Name) and generates a populated Excel file.
    public class RestApiDataSource : ICellsDataTable
    {
        private readonly List<Dictionary<string, JsonElement>> _rows;
        private readonly string[] _columns;
        private int _currentRow = -1;

        // ctor receives the API endpoint URL
        public RestApiDataSource(string apiUrl)
        {
            try
            {
                // Synchronously fetch data (for demo purposes)
                using var client = new HttpClient();
                var response = client.GetAsync(apiUrl).Result;
                response.EnsureSuccessStatusCode();
                var json = response.Content.ReadAsStringAsync().Result;

                // Expect the JSON to be an array of objects, e.g. [{ "Name":"John", "Age":30 }, ...]
                var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                if (root.ValueKind != JsonValueKind.Array)
                    throw new InvalidOperationException("API must return a JSON array.");

                _rows = new List<Dictionary<string, JsonElement>>();
                foreach (var element in root.EnumerateArray())
                {
                    var dict = new Dictionary<string, JsonElement>(StringComparer.OrdinalIgnoreCase);
                    foreach (var prop in element.EnumerateObject())
                    {
                        dict[prop.Name] = prop.Value;
                    }
                    _rows.Add(dict);
                }

                // Determine column names from the first row (if any)
                if (_rows.Count > 0)
                {
                    var first = _rows[0];
                    var cols = new List<string>(first.Keys);
                    _columns = cols.ToArray();
                }
                else
                {
                    _columns = Array.Empty<string>();
                }
            }
            catch (Exception ex)
            {
                // In case of any failure (e.g., network error, 404), fall back to empty data set
                Console.WriteLine($"Warning: Unable to retrieve data from API. {ex.Message}");
                _rows = new List<Dictionary<string, JsonElement>>();
                _columns = Array.Empty<string>();
            }
        }

        // ICellsDataTable members -------------------------------------------------
        public object this[int rowIndex, int columnIndex]
        {
            get
            {
                var colName = _columns[columnIndex];
                return GetValue(rowIndex, colName);
            }
        }

        public object this[int rowIndex] => _rows[rowIndex];

        public object this[string columnName]
        {
            get
            {
                if (_currentRow < 0 || _currentRow >= _rows.Count)
                    throw new IndexOutOfRangeException("Current row is out of range.");
                return GetValue(_currentRow, columnName);
            }
        }

        public int RowCount => _rows.Count;

        public int ColumnCount => _columns.Length;

        public int Count => _rows.Count;

        public string[] Columns => _columns;

        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        public bool Next()
        {
            _currentRow++;
            return _currentRow < _rows.Count;
        }

        // Helper to extract a .NET value from JsonElement
        private object GetValue(int rowIndex, string columnName)
        {
            if (rowIndex < 0 || rowIndex >= _rows.Count)
                throw new IndexOutOfRangeException("Row index out of range.");

            var row = _rows[rowIndex];
            if (!row.TryGetValue(columnName, out var jsonElem))
                return null;

            return jsonElem.ValueKind switch
            {
                JsonValueKind.String => jsonElem.GetString(),
                JsonValueKind.Number => jsonElem.TryGetInt64(out var l) ? (object)l :
                                        jsonElem.TryGetDouble(out var d) ? d : jsonElem.GetDecimal(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null => null,
                _ => jsonElem.GetRawText()
            };
        }
    }

    // Demonstration of using the custom data source with smart markers
    public static class SmartMarkerRestApiDemo
    {
        public static void Run()
        {
            try
            {
                // 1. Create a new workbook (template)
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Place smart markers that correspond to JSON fields returned by the API
                // Example assumes the API returns objects with fields: Name, Age, City
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Age");
                sheet.Cells["C1"].PutValue("City");

                // Row 2 will be populated by the smart marker engine
                sheet.Cells["A2"].PutValue("&=$ApiData.Name");
                sheet.Cells["B2"].PutValue("&=$ApiData.Age");
                sheet.Cells["C2"].PutValue("&=$ApiData.City");

                // 2. Create a WorkbookDesigner and bind the custom data source
                var designer = new WorkbookDesigner(workbook);

                // Replace with the actual REST endpoint that returns a JSON array
                const string apiUrl = "https://example.com/api/people";

                // Set the custom data source; the name "ApiData" matches the smart marker prefix
                designer.SetDataSource("ApiData", new RestApiDataSource(apiUrl));

                // 3. Process smart markers
                designer.Process();

                // 4. Save the populated workbook
                const string outputPath = "SmartMarkerFromRestApi.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during SmartMarker processing: {ex.Message}");
            }
        }
    }

    // Entry point
    internal class Program
    {
        private static void Main()
        {
            SmartMarkerRestApiDemo.Run();
            Console.WriteLine("Execution completed.");
        }
    }
}
