// Title: Import JSON array into an Excel worksheet with Aspose.Cells by converting to a DataTable in C#
// Common Searches: C# read JSON file and import into Excel using Aspose.Cells | Aspose.Cells ImportData from a DataTable built from JSON | convert JSON array to DataTable for Excel export with Aspose.Cells | display column names when importing JSON data into an Excel worksheet | adjust column widths after JSON import with Aspose.Cells
// Tags: Aspose.Cells import json data to worksheet | json to datatable conversion Aspose.Cells | ImportTableOptions column headers | auto fit columns Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using System.Data;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonImport
{
    // Parses a JSON array of objects and creates a DataTable where each property becomes a column.
    // The example reads a JSON file containing an array of objects, dynamically builds a DataTable with columns matching the JSON properties, and uses Aspose.Cells ImportData (with ImportTableOptions to show column names) to write the data to the first worksheet, auto‑fits the columns, and saves the workbook as an XLSX file.
    public class JsonDataTableMapper
    {
        // Returns a DataTable built from the supplied JSON string.
        // Expected JSON format:
        // [
        //   { "Id": 1, "Name": "Apple",  "Price": 0.5 },
        //   { "Id": 2, "Name": "Banana", "Price": 0.3 }
        // ]
        public DataTable GetDataTable(string json)
        {
            using JsonDocument doc = JsonDocument.Parse(json);
            JsonElement root = doc.RootElement;

            if (root.ValueKind != JsonValueKind.Array)
                throw new ArgumentException("Root element of JSON must be an array.");

            DataTable table = new DataTable("JsonData");
            bool columnsCreated = false;

            foreach (JsonElement element in root.EnumerateArray())
            {
                if (element.ValueKind != JsonValueKind.Object)
                    continue; // skip non‑object entries

                // Create columns on the first iteration based on the property names.
                if (!columnsCreated)
                {
                    foreach (JsonProperty prop in element.EnumerateObject())
                    {
                        Type columnType = GetClrType(prop.Value);
                        table.Columns.Add(prop.Name, columnType);
                    }
                    columnsCreated = true;
                }

                // Build a new DataRow.
                DataRow row = table.NewRow();
                foreach (JsonProperty prop in element.EnumerateObject())
                {
                    object value = GetClrValue(prop.Value);
                    row[prop.Name] = value ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }

            return table;
        }

        // Map a JsonElement to a CLR type for DataColumn creation.
        private static Type GetClrType(JsonElement element) =>
            element.ValueKind switch
            {
                JsonValueKind.Number => element.TryGetInt64(out _) ? typeof(long) : typeof(double),
                JsonValueKind.String => typeof(string),
                JsonValueKind.True or JsonValueKind.False => typeof(bool),
                JsonValueKind.Null => typeof(string),
                _ => typeof(string)
            };

        // Extract the actual CLR value from a JsonElement.
        private static object GetClrValue(JsonElement element) =>
            element.ValueKind switch
            {
                JsonValueKind.Number => element.TryGetInt64(out long l) ? (object)l :
                                         element.TryGetDouble(out double d) ? d : null,
                JsonValueKind.String => element.GetString(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null => null,
                _ => element.GetRawText()
            };
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the JSON file. Adjust as needed.
                const string jsonFilePath = "data.json";

                if (!File.Exists(jsonFilePath))
                {
                    Console.WriteLine($"JSON file not found: {jsonFilePath}");
                    return;
                }

                // Read the entire JSON content.
                string jsonContent = File.ReadAllText(jsonFilePath);

                // Create an instance of the custom mapper and obtain a DataTable.
                var mapper = new JsonDataTableMapper();
                DataTable jsonTable = mapper.GetDataTable(jsonContent);

                // Create a new workbook and get the first worksheet's cells.
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];
                var cells = worksheet.Cells;

                // Import the DataTable into the worksheet with column headers.
                var importOptions = new ImportTableOptions
                {
                    IsFieldNameShown = true   // Show column names as the first row.
                };
                cells.ImportData(jsonTable, 0, 0, importOptions);

                // Auto‑fit columns for better readability.
                worksheet.AutoFitColumns();

                // Save the workbook.
                const string outputPath = "JsonImported.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"JSON data imported successfully and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
