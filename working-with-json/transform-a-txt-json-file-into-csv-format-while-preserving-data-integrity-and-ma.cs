using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace JsonToCsvConverter
{
    class Program
    {
        static void Main()
        {
            // Paths for input JSON file and output CSV file
            string jsonFilePath = "input.json";
            string csvFilePath = "output.csv";

            // Read the entire JSON content
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Parse JSON; support array or single object
            List<JsonElement> elements = new List<JsonElement>();
            try
            {
                using JsonDocument doc = JsonDocument.Parse(jsonContent);
                JsonElement root = doc.RootElement;

                if (root.ValueKind == JsonValueKind.Array)
                {
                    foreach (JsonElement el in root.EnumerateArray())
                        elements.Add(el);
                }
                else if (root.ValueKind == JsonValueKind.Object)
                {
                    elements.Add(root);
                }
                else
                {
                    Console.WriteLine("JSON root element is neither an array nor an object.");
                    return;
                }
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"Failed to parse JSON: {ex.Message}");
                return;
            }

            if (elements.Count == 0)
            {
                Console.WriteLine("No data found in JSON.");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Write column headers from the first object
            int colIndex = 0;
            foreach (JsonProperty prop in elements[0].EnumerateObject())
            {
                cells[0, colIndex].PutValue(prop.Name);
                colIndex++;
            }

            // Write data rows
            int currentRow = 0;
            foreach (JsonElement element in elements)
            {
                colIndex = 0;
                foreach (JsonProperty prop in element.EnumerateObject())
                {
                    switch (prop.Value.ValueKind)
                    {
                        case JsonValueKind.Number:
                            if (prop.Value.TryGetInt64(out long l))
                                cells[currentRow + 1, colIndex].PutValue(l);
                            else if (prop.Value.TryGetDouble(out double d))
                                cells[currentRow + 1, colIndex].PutValue(d);
                            break;
                        case JsonValueKind.True:
                        case JsonValueKind.False:
                            cells[currentRow + 1, colIndex].PutValue(prop.Value.GetBoolean());
                            break;
                        case JsonValueKind.String:
                            cells[currentRow + 1, colIndex].PutValue(prop.Value.GetString());
                            break;
                        default:
                            cells[currentRow + 1, colIndex].PutValue(prop.Value.GetRawText());
                            break;
                    }
                    colIndex++;
                }
                currentRow++;
            }

            // Save the workbook as CSV using TxtSaveOptions to control separator
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = ',',
                Encoding = Encoding.UTF8,
                KeepSeparatorsForBlankRow = false,
                TrimLeadingBlankRowAndColumn = true
            };

            workbook.Save(csvFilePath, saveOptions);

            Console.WriteLine($"JSON data successfully converted to CSV at '{csvFilePath}'.");
        }
    }
}