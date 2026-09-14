// Title: Use a custom ICellsDataTable to load JSON from a REST API and fill Aspose.Cells smart markers in C#
// AI Prompts: Write a C# class that implements ICellsDataTable to transform a JSON array retrieved with HttpClient into a tabular source for Aspose.Cells smart markers. | Show how to attach the custom ICellsDataTable to a WorkbookDesigner, place smart markers such as &=$Products.Name, process them, and save the resulting Excel workbook. | Extend the ApiDataSource to handle nested JSON objects and map them to hierarchical smart markers in an Excel template. | Add robust error handling for the HTTP request and fallback to sample JSON while still using the custom data source.
// Common Searches: how to bind a REST API JSON response to Aspose.Cells smart markers using ICellsDataTable in C# | example of custom data source for smart markers with Aspose.Cells and HttpClient | populate Excel smart markers from dynamic JSON array in Aspose.Cells | C# code to implement ICellsDataTable for API data and use WorkbookDesigner
// Tags: custom ICellsDataTable JSON API integration | Aspose.Cells smart markers from REST endpoint | WorkbookDesigner bind API data source | populate Excel with smart markers using HttpClient | dynamic JSON data source for Aspose.Cells

using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Custom data source that implements ICellsDataTable.
    // It fetches JSON data from a REST API and exposes it in a tabular form.
    // The sample defines an ApiDataSource class that implements ICellsDataTable, parses a JSON array returned by an HTTP GET call (or a fallback sample), and exposes rows and columns for smart markers. In Main it fetches JSON via HttpClient, creates a Workbook, inserts smart markers (&=$Products.Id, Name, Price), binds the custom data source to the "Products" marker name using WorkbookDesigner, processes the markers, and saves the populated workbook as SmartMarkerFromApi.xlsx.
    public class ApiDataSource : ICellsDataTable
    {
        private readonly List<Dictionary<string, object>> _rows = new List<Dictionary<string, object>>();
        private readonly string[] _columns;
        private int _currentRow = -1;

        // Constructor accepts the JSON array string and extracts column names.
        public ApiDataSource(string jsonArray)
        {
            // Parse the JSON array.
            var jsonDoc = JsonDocument.Parse(jsonArray);
            foreach (var element in jsonDoc.RootElement.EnumerateArray())
            {
                var dict = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                foreach (var property in element.EnumerateObject())
                {
                    // Store primitive values as strings; otherwise keep raw JSON.
                    dict[property.Name] = property.Value.ValueKind == JsonValueKind.String
                        ? property.Value.GetString()
                        : (object)property.Value.GetRawText();
                }
                _rows.Add(dict);
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

        // Indexer for row/column access (zero‑based).
        public object this[int rowIndex, int columnIndex] => _rows[rowIndex][_columns[columnIndex]];

        // Indexer for row access (returns the whole row object).
        public object this[int rowIndex] => _rows[rowIndex];

        // Indexer for column name access (uses the current row pointer).
        public object this[string columnName] => _rows[_currentRow][columnName];

        public int RowCount => _rows.Count;
        public int ColumnCount => _columns.Length;
        public int Count => _rows.Count;
        public string[] Columns => _columns;

        // Resets the internal pointer before iteration.
        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        // Moves to the next row; returns false when no more rows.
        public bool Next()
        {
            _currentRow++;
            return _currentRow < _rows.Count;
        }
    }

    class Program
    {
        static void Main()
        {
            const string apiUrl = "https://example.com/api/products";

            string jsonResponse;

            // Attempt to retrieve JSON data from the API; fall back to sample data on failure.
            try
            {
                using var httpClient = new HttpClient();
                jsonResponse = httpClient.GetStringAsync(apiUrl).Result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to fetch data from API: {ex.Message}");
                // Sample JSON array to demonstrate functionality.
                jsonResponse = @"[
                    { ""Id"": ""1"", ""Name"": ""Product A"", ""Price"": ""10.99"" },
                    { ""Id"": ""2"", ""Name"": ""Product B"", ""Price"": ""15.49"" },
                    { ""Id"": ""3"", ""Name"": ""Product C"", ""Price"": ""7.25"" }
                ]";
            }

            try
            {
                // Create a new workbook (or load a template if you have one).
                var workbook = new Workbook();

                // Access the first worksheet.
                var sheet = workbook.Worksheets[0];

                // Place smart markers that correspond to the JSON fields.
                sheet.Cells["A1"].PutValue("&=$Products.Id");
                sheet.Cells["B1"].PutValue("&=$Products.Name");
                sheet.Cells["C1"].PutValue("&=$Products.Price");

                // Create the workbook designer and assign the workbook.
                var designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Wrap the JSON response in the custom ICellsDataTable implementation.
                var apiDataSource = new ApiDataSource(jsonResponse);

                // Bind the custom data source to the smart marker name "Products".
                designer.SetDataSource("Products", apiDataSource);

                // Process the smart markers and populate the worksheet.
                designer.Process();

                // Save the result.
                const string outputPath = "SmartMarkerFromApi.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }
}
