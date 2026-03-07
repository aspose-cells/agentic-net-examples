using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace OxpsJsonToExcel
{
    class Program
    {
        static void Main()
        {
            string jsonFilePath = "input.json";
            string excelFilePath = "output.xlsx";

            if (!File.Exists(jsonFilePath))
            {
                Console.WriteLine($"JSON file not found: {jsonFilePath}");
                return;
            }

            string jsonContent = File.ReadAllText(jsonFilePath);

            var records = new List<Dictionary<string, JsonElement>>();

            try
            {
                using (JsonDocument doc = JsonDocument.Parse(jsonContent))
                {
                    if (doc.RootElement.ValueKind == JsonValueKind.Array)
                    {
                        foreach (JsonElement element in doc.RootElement.EnumerateArray())
                        {
                            if (element.ValueKind == JsonValueKind.Object)
                            {
                                var dict = new Dictionary<string, JsonElement>();
                                foreach (JsonProperty prop in element.EnumerateObject())
                                {
                                    dict[prop.Name] = prop.Value;
                                }
                                records.Add(dict);
                            }
                        }
                    }
                    else if (doc.RootElement.ValueKind == JsonValueKind.Object)
                    {
                        var dict = new Dictionary<string, JsonElement>();
                        foreach (JsonProperty prop in doc.RootElement.EnumerateObject())
                        {
                            dict[prop.Name] = prop.Value;
                        }
                        records.Add(dict);
                    }
                }
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"Failed to parse JSON: {ex.Message}");
                return;
            }

            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            if (records.Count > 0)
            {
                int colIndex = 0;
                foreach (var header in records[0].Keys)
                {
                    sheet.Cells[0, colIndex].PutValue(header);
                    colIndex++;
                }

                for (int rowIndex = 0; rowIndex < records.Count; rowIndex++)
                {
                    var record = records[rowIndex];
                    colIndex = 0;
                    foreach (var value in record.Values)
                    {
                        switch (value.ValueKind)
                        {
                            case JsonValueKind.String:
                                sheet.Cells[rowIndex + 1, colIndex].PutValue(value.GetString());
                                break;
                            case JsonValueKind.Number:
                                if (value.TryGetInt64(out long l))
                                    sheet.Cells[rowIndex + 1, colIndex].PutValue(l);
                                else if (value.TryGetDouble(out double d))
                                    sheet.Cells[rowIndex + 1, colIndex].PutValue(d);
                                break;
                            case JsonValueKind.True:
                            case JsonValueKind.False:
                                sheet.Cells[rowIndex + 1, colIndex].PutValue(value.GetBoolean());
                                break;
                            case JsonValueKind.Null:
                                sheet.Cells[rowIndex + 1, colIndex].PutValue(string.Empty);
                                break;
                            default:
                                sheet.Cells[rowIndex + 1, colIndex].PutValue(value.GetRawText());
                                break;
                        }
                        colIndex++;
                    }
                }
            }

            workbook.Save(excelFilePath, SaveFormat.Xlsx);
            Console.WriteLine($"Conversion completed. Excel file saved to '{excelFilePath}'.");
        }
    }
}