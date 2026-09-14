// Title: Export an Excel worksheet to an indented JSON array with column headers as object keys using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, reads the first row as field names, converts each subsequent row into a dictionary while preserving the original cell data types, and returns a formatted JSON string. | Update the example to write the resulting JSON to a file named "output.json" on disk instead of printing it to the console. | Add logic to exclude columns whose header is empty or starts with an underscore, so those columns are omitted from the JSON output.
// Common Searches: how to convert an Excel sheet to JSON with column names as keys using Aspose.Cells C# | Aspose.Cells export worksheet data to indented JSON string preserving data types | C# read Excel first row as headers and serialize rows to JSON array | ignore empty rows when exporting Excel to JSON with Aspose.Cells | exclude columns with empty or underscore headers in Excel to JSON conversion .NET
// Tags: Aspose.Cells export worksheet to JSON | C# serialize Excel rows to JSON objects | preserve Excel cell data types in JSON | ignore empty rows in Excel JSON export | filter columns by header name in Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Text.Json;

// The sample loads an Excel workbook with Aspose.Cells, extracts column names from the first row, builds a list of dictionaries for each subsequent row while keeping original cell types, skips completely empty rows, serializes the collection to an indented JSON string, and outputs the JSON (or optionally writes it to a file).
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can change the index as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Access the cells collection
        Cells cells = sheet.Cells;

        // Determine the used range boundaries
        int maxRow = cells.MaxDataRow;      // zero‑based index of last row with data
        int maxCol = cells.MaxDataColumn;   // zero‑based index of last column with data

        // -----------------------------------------------------------------
        // Read column names from the first row (row index 0)
        // -----------------------------------------------------------------
        List<string> headers = new List<string>();
        for (int col = 0; col <= maxCol; col++)
        {
            string header = cells[0, col].StringValue;
            // If a header cell is empty, assign a default name
            if (string.IsNullOrEmpty(header))
                header = $"Column{col}";
            headers.Add(header);
        }

        // -----------------------------------------------------------------
        // Build a list of dictionaries where each dictionary represents a row
        // -----------------------------------------------------------------
        List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();

        // Start from row 1 because row 0 contains headers
        for (int row = 1; row <= maxRow; row++)
        {
            var dict = new Dictionary<string, object>();
            bool isEmptyRow = true;

            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                object value = null;

                // Preserve the original data type where possible
                switch (cell.Type)
                {
                    case CellValueType.IsNumeric:
                        value = cell.DoubleValue;
                        break;
                    case CellValueType.IsString:
                        value = cell.StringValue;
                        break;
                    case CellValueType.IsBool:
                        value = cell.BoolValue;
                        break;
                    case CellValueType.IsDateTime:
                        value = cell.DateTimeValue;
                        break;
                    default:
                        value = cell.Value;
                        break;
                }

                // Detect if the row has any non‑empty cell
                if (value != null && !string.IsNullOrEmpty(value.ToString()))
                    isEmptyRow = false;

                dict[headers[col]] = value;
            }

            // Skip rows that are completely empty
            if (!isEmptyRow)
                rows.Add(dict);
        }

        // -----------------------------------------------------------------
        // Serialize the list to a JSON string, using indented formatting
        // -----------------------------------------------------------------
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string json = JsonSerializer.Serialize(rows, jsonOptions);

        // Output the JSON string (you could also write it to a file)
        Console.WriteLine(json);
    }
}
