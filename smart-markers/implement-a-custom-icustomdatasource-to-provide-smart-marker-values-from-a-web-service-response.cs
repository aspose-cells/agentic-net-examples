using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsCustomDataSourceDemo
{
    // Custom data source that implements ICellsDataTable.
    // It retrieves JSON data from a web service and presents it in a tabular form
    // that can be consumed by Aspose.Cells smart markers.
    public class WebServiceDataSource : ICellsDataTable
    {
        private readonly List<Dictionary<string, object>> _rows = new List<Dictionary<string, object>>();
        private readonly string[] _columns;
        private int _currentRow = -1;

        // Constructor fetches data from the specified URL.
        public WebServiceDataSource(string requestUrl)
        {
            try
            {
                // Synchronous call for simplicity; in production use async/await.
                using (HttpClient client = new HttpClient())
                {
                    string json = client.GetStringAsync(requestUrl).Result;
                    // Expecting a JSON array of objects, e.g. [{ "Name":"A", "Price":10 }, ...]
                    JsonDocument doc = JsonDocument.Parse(json);
                    foreach (JsonElement element in doc.RootElement.EnumerateArray())
                    {
                        var dict = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                        foreach (JsonProperty prop in element.EnumerateObject())
                        {
                            // Store primitive values; complex types can be extended as needed.
                            dict[prop.Name] = prop.Value.ValueKind switch
                            {
                                JsonValueKind.String => prop.Value.GetString(),
                                JsonValueKind.Number => prop.Value.GetDouble(),
                                JsonValueKind.True => true,
                                JsonValueKind.False => false,
                                _ => prop.Value.GetRawText()
                            };
                        }
                        _rows.Add(dict);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the error and continue with an empty data set.
                Console.WriteLine($"Error fetching data from web service: {ex.Message}");
            }

            // Determine column names from the first row (if any).
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

        // Indexer by row and column index.
        public object this[int rowIndex, int columnIndex] => _rows[rowIndex][_columns[columnIndex]];

        // Indexer by row index only – returns the whole row as an object.
        public object this[int rowIndex] => _rows[rowIndex];

        // Indexer by column name – works on the current row (set by Next()).
        public object this[string columnName] => _rows[_currentRow][columnName];

        // Total number of rows.
        public int RowCount => _rows.Count;

        // Total number of columns.
        public int ColumnCount => _columns.Length;

        // Alias for RowCount (required by ICellsDataTable).
        public int Count => RowCount;

        // Column names.
        public string[] Columns => _columns;

        // Reset enumeration.
        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        // Move to the next row; returns false when no more rows.
        public bool Next()
        {
            _currentRow++;
            return _currentRow < RowCount;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string resultPath = "Result.xlsx";

                // Ensure the template file exists; create a simple one if missing.
                if (!File.Exists(templatePath))
                {
                    var wb = new Workbook();
                    var ws = wb.Worksheets[0];
                    ws.Cells["A1"].PutValue("&=$Products.Name");
                    ws.Cells["B1"].PutValue("&=$Products.Price");
                    wb.Save(templatePath);
                    Console.WriteLine($"Template file not found. Created a default template at '{templatePath}'.");
                }

                // Load the workbook that contains smart markers.
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = new Workbook(templatePath)
                };

                // Create and assign the custom data source.
                // The URL should return a JSON array of product objects.
                string serviceUrl = "https://example.com/api/products"; // replace with actual service URL
                designer.SetDataSource("Products", new WebServiceDataSource(serviceUrl));

                // Process smart markers using the custom data source.
                designer.Process();

                // Save the populated workbook.
                designer.Workbook.Save(resultPath);
                Console.WriteLine($"Workbook processed and saved to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}