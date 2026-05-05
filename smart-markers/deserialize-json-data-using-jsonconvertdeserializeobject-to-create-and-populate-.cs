using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsJsonDemo
{
    class Program
    {
        static void Main()
        {
            // Sample JSON string (array of objects)
            string json = @"[
                { ""Name"": ""John"", ""Age"": 30, ""City"": ""New York"" },
                { ""Name"": ""Alice"", ""Age"": 25, ""City"": ""London"" },
                { ""Name"": ""Bob"", ""Age"": 28, ""City"": ""Paris"" }
            ]";

            // Deserialize JSON into a list of dictionaries using System.Text.Json
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            List<Dictionary<string, JsonElement>> records = JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(json, options);

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            if (records != null && records.Count > 0)
            {
                // Write header row using keys from the first record
                int colIndex = 0;
                foreach (string header in records[0].Keys)
                {
                    cells[0, colIndex].PutValue(header);
                    colIndex++;
                }

                // Write data rows
                for (int row = 0; row < records.Count; row++)
                {
                    colIndex = 0;
                    foreach (var kvp in records[row])
                    {
                        object value = GetJsonElementValue(kvp.Value);
                        cells[row + 1, colIndex].PutValue(value);
                        colIndex++;
                    }
                }
            }

            // Save the workbook to an XLSX file
            workbook.Save("OutputFromJson.xlsx", SaveFormat.Xlsx);
        }

        private static object GetJsonElementValue(JsonElement element)
        {
            return element.ValueKind switch
            {
                JsonValueKind.String => element.GetString(),
                JsonValueKind.Number => element.TryGetInt64(out long l) ? (object)l : element.GetDouble(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null => null,
                _ => element.GetRawText()
            };
        }
    }
}