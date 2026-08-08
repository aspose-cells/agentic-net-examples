// Title: Import JSON into Excel with a custom ICellsDataTable – Aspose.Cells for .NET
// Description: Parse a JSON array, implement ICellsDataTable, and use Worksheet.Cells.ImportData (with ImportTableOptions) to write the data and headers to a new XLSX file.
// Keywords: Aspose.Cells | ICellsDataTable | JSON to Excel | C# .NET | ImportData | ImportTableOptions | Excel automation | worksheet import | data mapping | JSON parsing
// Common Searches: Aspose.Cells import JSON array | Implement ICellsDataTable in C# | Map JSON fields to Excel columns | ImportData with headers Aspose.Cells | Insert rows while importing data | C# convert JSON to XLSX | Custom data table for Aspose.Cells
// Developer Intent: Build an ICellsDataTable that reads JSON and import the resulting rows into an Excel sheet.
// Use Cases: Convert a JSON list of objects into an Excel workbook with column headers. | Import large JSON datasets into an existing worksheet without overwriting data by enabling InsertRows. | Reuse the JsonCellsDataTable class for different JSON structures by adjusting column extraction logic.
// AI Prompts: Create a JsonCellsDataTable that flattens nested JSON objects into dot‑separated columns. | Show how to configure ImportTableOptions to omit the header row during import. | Provide streaming code to read a massive JSON file and feed rows to ICellsDataTable on the fly.

using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonImport
{
    // Custom implementation of ICellsDataTable that holds JSON data
    // Parse a JSON array, implement ICellsDataTable, and use Worksheet.Cells.ImportData (with ImportTableOptions) to write the data and headers to a new XLSX file.
    public class JsonCellsDataTable : ICellsDataTable
    {
        private readonly List<string> _columns;
        private readonly List<Dictionary<string, object>> _rows;
        private int _cursor = -1; // Position before the first row

        public JsonCellsDataTable(List<string> columns, List<Dictionary<string, object>> rows)
        {
            _columns = columns;
            _rows = rows;
        }

        // Returns column names
        public string[] Columns => _columns.ToArray();

        // Returns number of records
        public int Count => _rows.Count;

        // Indexer by column index
        public object this[int columnIndex]
        {
            get
            {
                if (_cursor < 0 || _cursor >= _rows.Count)
                    throw new InvalidOperationException("Cursor is not positioned on a valid row.");

                string colName = _columns[columnIndex];
                return _rows[_cursor].TryGetValue(colName, out var value) ? value : null;
            }
        }

        // Indexer by column name
        public object this[string columnName]
        {
            get
            {
                if (_cursor < 0 || _cursor >= _rows.Count)
                    throw new InvalidOperationException("Cursor is not positioned on a valid row.");

                return _rows[_cursor].TryGetValue(columnName, out var value) ? value : null;
            }
        }

        // Move cursor to before the first row
        public void BeforeFirst()
        {
            _cursor = -1;
        }

        // Move cursor to the next row; returns false if no more rows
        public bool Next()
        {
            if (_cursor + 1 < _rows.Count)
            {
                _cursor++;
                return true;
            }
            return false;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Sample JSON array
            string json = @"[
                { ""Name"": ""Alice"", ""Age"": 30, ""City"": ""New York"" },
                { ""Name"": ""Bob"",   ""Age"": 25, ""City"": ""Los Angeles"" },
                { ""Name"": ""Charlie"", ""Age"": 35, ""City"": ""Chicago"" }
            ]";

            // Parse JSON and build column list + rows
            var columns = new List<string>();
            var rows = new List<Dictionary<string, object>>();

            using (JsonDocument doc = JsonDocument.Parse(json))
            {
                JsonElement root = doc.RootElement;
                if (root.ValueKind != JsonValueKind.Array)
                    throw new InvalidOperationException("Root JSON element must be an array.");

                foreach (JsonElement element in root.EnumerateArray())
                {
                    var dict = new Dictionary<string, object>();
                    foreach (JsonProperty prop in element.EnumerateObject())
                    {
                        // Capture column names from the first element
                        if (columns.Count == 0 && !columns.Contains(prop.Name))
                            columns.Add(prop.Name);

                        // Store value (handle different JSON value kinds)
                        object value = prop.Value.ValueKind switch
                        {
                            JsonValueKind.String => prop.Value.GetString(),
                            JsonValueKind.Number => prop.Value.TryGetInt64(out long l) ? (object)l :
                                                    prop.Value.TryGetDouble(out double d) ? d : null,
                            JsonValueKind.True => true,
                            JsonValueKind.False => false,
                            JsonValueKind.Null => null,
                            _ => prop.Value.GetRawText()
                        };
                        dict[prop.Name] = value;
                    }
                    rows.Add(dict);
                }
            }

            // Create custom ICellsDataTable from parsed JSON
            ICellsDataTable jsonTable = new JsonCellsDataTable(columns, rows);

            // Create workbook and import the data table
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Import with field names shown in the first row
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,
                InsertRows = true
            };
            sheet.Cells.ImportData(jsonTable, 0, 0, importOptions);

            // Save the workbook
            workbook.Save("JsonImportOutput.xlsx");
        }
    }
}
