// Title: Import JSON Array into Excel with a Custom ICellsDataTable – Aspose.Cells for .NET
// Description: Demonstrates how to implement ICellsDataTable to read a JSON array, map its fields to worksheet columns, and import the data into an Excel workbook using Aspose.Cells' ImportData method. The example uses System.Text.Json, dictionary‑based rows, and saves the result as an XLSX file.
// Keywords: Aspose.Cells JSON import | ICellsDataTable example | C# import JSON to Excel | System.Text.Json Aspose.Cells | ImportData custom data table | Excel workbook from JSON | .NET Excel automation
// Common Searches: Aspose.Cells import JSON array C# | How to use ICellsDataTable with JSON | Map JSON fields to Excel columns Aspose | Custom data table for ImportData Aspose.Cells | C# example reading JSON into Excel worksheet
// Developer Intent: Create a reusable ICellsDataTable that parses a JSON array and feeds it directly into an Aspose.Cells worksheet.
// Use Cases: Generate Excel reports from API responses that return JSON collections. | Convert flat JSON files (e.g., employee or product lists) into spreadsheets without defining a strong‑typed model. | Build a generic JSON‑to‑Excel utility that can be applied across multiple projects or worksheets.
// AI Prompts: Write C# code that reads a JSON file, constructs a JsonCellsDataTable, and imports it into an Aspose.Cells workbook with custom column ordering. | Show how to extend JsonCellsDataTable to handle nested objects or arrays for hierarchical Excel export. | Create unit tests for JsonCellsDataTable covering Next, BeforeFirst, and indexer access by column name and index.

using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsJsonImport
{
    // Custom implementation of ICellsDataTable that reads data from a JSON array
    // Demonstrates how to implement ICellsDataTable to read a JSON array, map its fields to worksheet columns, and import the data into an Excel workbook using Aspose.Cells' ImportData method. The example uses System.Text.Json, dictionary‑based rows, and saves the result as an XLSX file.
    public class JsonCellsDataTable : ICellsDataTable
    {
        private readonly List<Dictionary<string, object>> _rows;
        private readonly string[] _columns;
        private int _currentRow = -1; // Position before the first row

        public JsonCellsDataTable(List<Dictionary<string, object>> rows, string[] columns)
        {
            _rows = rows;
            _columns = columns;
        }

        // Returns the column names
        public string[] Columns => _columns;

        // Returns the number of records; -1 if unknown (not used here)
        public int Count => _rows.Count;

        // Indexer by column index
        public object this[int columnIndex]
        {
            get
            {
                if (_currentRow < 0 || _currentRow >= _rows.Count)
                    throw new IndexOutOfRangeException("Current row is out of range.");

                string colName = _columns[columnIndex];
                return _rows[_currentRow].TryGetValue(colName, out var value) ? value : null;
            }
        }

        // Indexer by column name
        public object this[string columnName]
        {
            get
            {
                if (_currentRow < 0 || _currentRow >= _rows.Count)
                    throw new IndexOutOfRangeException("Current row is out of range.");

                return _rows[_currentRow].TryGetValue(columnName, out var value) ? value : null;
            }
        }

        // Move cursor to before the first row
        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        // Move cursor to the next row; returns false if no more rows
        public bool Next()
        {
            if (_currentRow + 1 < _rows.Count)
            {
                _currentRow++;
                return true;
            }
            return false;
        }
    }

    public class JsonImportDemo
    {
        public static void Run()
        {
            // Sample JSON array (each object represents a record)
            string jsonInput = @"[
                { ""Name"": ""Alice"", ""Age"": 30, ""Country"": ""USA"" },
                { ""Name"": ""Bob"",   ""Age"": 25, ""Country"": ""UK"" },
                { ""Name"": ""Carol"", ""Age"": 28, ""Country"": ""Canada"" }
            ]";

            // Parse JSON and build rows + column list
            var rows = new List<Dictionary<string, object>>();
            string[] columns = null;

            using (JsonDocument doc = JsonDocument.Parse(jsonInput))
            {
                foreach (JsonElement element in doc.RootElement.EnumerateArray())
                {
                    var dict = new Dictionary<string, object>();
                    foreach (JsonProperty prop in element.EnumerateObject())
                    {
                        // Store primitive values; complex types can be extended as needed
                        dict[prop.Name] = prop.Value.ValueKind switch
                        {
                            JsonValueKind.Number when prop.Value.TryGetInt32(out int i) => i,
                            JsonValueKind.Number when prop.Value.TryGetDouble(out double d) => d,
                            JsonValueKind.String => prop.Value.GetString(),
                            JsonValueKind.True => true,
                            JsonValueKind.False => false,
                            _ => prop.Value.GetRawText()
                        };
                    }
                    rows.Add(dict);

                    // Capture column names from the first object
                    if (columns == null)
                    {
                        var colList = new List<string>();
                        foreach (var kvp in dict)
                            colList.Add(kvp.Key);
                        columns = colList.ToArray();
                    }
                }
            }

            // Create the custom ICellsDataTable instance
            ICellsDataTable dataTable = new JsonCellsDataTable(rows, columns);

            // Create workbook and import the data table into the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.ImportData(dataTable, 0, 0, new ImportTableOptions());

            // Save the workbook (using the provided lifecycle rule)
            workbook.Save("JsonImported.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            JsonImportDemo.Run();
        }
    }
}
