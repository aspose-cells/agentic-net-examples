// Title: How to import a JSON array into an Excel worksheet using a custom ICellsDataTable implementation with Aspose.Cells for .NET
// AI Prompts: Create a C# class that implements ICellsDataTable to read rows from a JSON string and expose column values for Aspose.Cells import. | Demonstrate using Cells.ImportData together with a custom ICellsDataTable and ImportTableOptions to write JSON data into a worksheet. | Update the ICellsDataTable implementation to correctly handle nulls and the various JSON primitive types during Excel import.
// Common Searches: aspnet import json array into excel using aspose.cells ICellsDataTable | c# ICellsDataTable example for reading JSON data with Aspose.Cells | map json fields to Excel columns using Aspose.Cells ImportData method | sample code to import JSON array into .xlsx using Aspose.Cells .NET
// Tags: ICellsDataTable JSON source handling | Aspose.Cells import JSON array to XLSX | Excel worksheet population from JSON data | primitive JSON type conversion in Aspose.Cells | JSON deserialization to dictionary for Aspose.Cells

using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsJsonImport
{
    // Custom ICellsDataTable implementation that reads data from a JSON array.
    // Implements ICellsDataTable by deserializing a JSON array into a list of dictionaries, exposing column names and row values so that Cells.ImportData can write the data into an Excel worksheet.
    public class JsonCellsDataTable : ICellsDataTable
    {
        private readonly List<Dictionary<string, object>> _rows;
        private readonly string[] _columns;
        private int _cursor = -1; // Position before the first row.

        public JsonCellsDataTable(string json)
        {
            // Deserialize JSON array of objects into a list of dictionaries.
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            var elements = JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(json, options);
            _rows = new List<Dictionary<string, object>>();

            if (elements != null && elements.Count > 0)
            {
                // Determine column names from the first element.
                var first = elements[0];
                var columnList = new List<string>(first.Keys);
                _columns = columnList.ToArray();

                // Convert each JsonElement to a .NET primitive and store rows.
                foreach (var dict in elements)
                {
                    var row = new Dictionary<string, object>();
                    foreach (var col in _columns)
                    {
                        if (dict.TryGetValue(col, out JsonElement je))
                        {
                            object value = je.ValueKind switch
                            {
                                JsonValueKind.String => je.GetString(),
                                JsonValueKind.Number => je.TryGetInt64(out long l) ? (object)l : je.GetDouble(),
                                JsonValueKind.True => true,
                                JsonValueKind.False => false,
                                JsonValueKind.Null => null,
                                JsonValueKind.Undefined => null,
                                JsonValueKind.Object => je.GetRawText(),
                                JsonValueKind.Array => je.GetRawText(),
                                _ => je.GetRawText()
                            };
                            row[col] = value;
                        }
                        else
                        {
                            row[col] = null;
                        }
                    }
                    _rows.Add(row);
                }
            }
            else
            {
                _columns = Array.Empty<string>();
            }
        }

        // ICellsDataTable members.

        public string[] Columns => _columns;

        public int Count => _rows.Count;

        // Indexer by column index.
        public object this[int columnIndex]
        {
            get
            {
                if (_cursor < 0 || _cursor >= _rows.Count)
                    throw new IndexOutOfRangeException("Cursor is not positioned on a valid row.");
                string colName = _columns[columnIndex];
                return _rows[_cursor][colName];
            }
        }

        // Indexer by column name.
        public object this[string columnName]
        {
            get
            {
                if (_cursor < 0 || _cursor >= _rows.Count)
                    throw new IndexOutOfRangeException("Cursor is not positioned on a valid row.");
                return _rows[_cursor].TryGetValue(columnName, out var val) ? val : null;
            }
        }

        public void BeforeFirst()
        {
            _cursor = -1;
        }

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

    class Program
    {
        static void Main()
        {
            // Sample JSON array to import.
            string json = @"[
                { ""Name"": ""Alice"",   ""Age"": 30, ""IsMember"": true },
                { ""Name"": ""Bob"",     ""Age"": 25, ""IsMember"": false },
                { ""Name"": ""Charlie"", ""Age"": 35, ""IsMember"": true }
            ]";

            // Create a new workbook and get the cells collection.
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Build a custom ICellsDataTable from the JSON string.
            ICellsDataTable jsonTable = new JsonCellsDataTable(json);

            // Import the data table into the worksheet starting at cell A1 (row 0, column 0).
            cells.ImportData(jsonTable, 0, 0, new ImportTableOptions());

            // Save the workbook.
            workbook.Save("JsonImportResult.xlsx");
        }
    }
}
