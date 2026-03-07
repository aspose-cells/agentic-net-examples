using System;
using System.IO;
using System.Text.Json;
using System.Collections.Generic;
using Aspose.Cells;

namespace OdsJsonToExcelConversion
{
    class Program
    {
        static void Main()
        {
            // Paths for the source JSON file (containing ODS‑like data) and the target Excel file
            string jsonFilePath = "source.json";
            string excelFilePath = "converted.xlsx";

            // Read the entire JSON content
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Parse the JSON; expecting an array of objects where each object represents a row
            JsonDocument doc = JsonDocument.Parse(jsonContent);
            JsonElement root = doc.RootElement;

            if (root.ValueKind != JsonValueKind.Array)
            {
                Console.WriteLine("The JSON file does not contain a top‑level array.");
                return;
            }

            // Convert JSON elements to a list of dictionaries for easier handling
            List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
            foreach (JsonElement element in root.EnumerateArray())
            {
                if (element.ValueKind != JsonValueKind.Object) continue;

                var dict = new Dictionary<string, object>();
                foreach (JsonProperty prop in element.EnumerateObject())
                {
                    // Store the raw JSON value; Aspose.Cells will handle basic types automatically
                    dict[prop.Name] = prop.Value.ValueKind switch
                    {
                        JsonValueKind.String => prop.Value.GetString(),
                        JsonValueKind.Number => prop.Value.TryGetInt64(out long l) ? (object)l :
                                                prop.Value.TryGetDouble(out double d) ? d : prop.Value.GetRawText(),
                        JsonValueKind.True => true,
                        JsonValueKind.False => false,
                        JsonValueKind.Null => null,
                        _ => prop.Value.GetRawText()
                    };
                }
                rows.Add(dict);
            }

            if (rows.Count == 0)
            {
                Console.WriteLine("No data rows found in the JSON file.");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Write header row (use keys from the first dictionary)
            int colIndex = 0;
            foreach (string header in rows[0].Keys)
            {
                sheet.Cells[0, colIndex].PutValue(header);
                colIndex++;
            }

            // Write data rows
            for (int i = 0; i < rows.Count; i++)
            {
                var rowDict = rows[i];
                colIndex = 0;
                foreach (var value in rowDict.Values)
                {
                    sheet.Cells[i + 1, colIndex].PutValue(value);
                    colIndex++;
                }
            }

            // Save the workbook as an Excel file (XLSX)
            workbook.Save(excelFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed. Excel file saved to '{excelFilePath}'.");
        }
    }
}