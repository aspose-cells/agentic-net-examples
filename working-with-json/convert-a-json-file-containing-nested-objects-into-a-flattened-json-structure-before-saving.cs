// Title: Flatten nested JSON exported from an Excel workbook using Aspose.Cells in C# and save it as a .json file
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, converts the workbook to a JSON string, recursively flattens the JSON into dot‑separated paths, and writes the flattened result to a .json file. | Write a C# recursive method that takes a System.Text.Json JsonElement and populates a Dictionary<string, object?> with full property paths, handling objects, arrays, and primitive values. | Adjust the flattening routine to use underscore separators for property paths while preserving array index notation, and output the result with System.Text.Json.
// Common Searches: c# flatten json produced by Aspose.Cells export workbook to json | how to convert nested json from excel to flat key‑value pairs using Aspose.Cells | recursive json flattening with dot notation in .net core | export excel to json and flatten structure with Aspose.Cells library | flatten json arrays to indexed keys in c# using System.Text.Json
// Tags: Aspose.Cells export workbook to JSON C# | C# recursive JSON flattening | dot‑separated property paths for flattened JSON | flatten JSON arrays with index keys C# | write flattened JSON file using System.Text.Json

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using Aspose.Cells;

// The program loads an Excel file via Aspose.Cells, saves the workbook as a JSON string, recursively flattens the nested JSON into dot‑separated keys (including array indices), and writes the flattened JSON to an output file.
class Program
{
    static void Main()
    {
        const string inputExcelPath = "input.xlsx";
        const string outputJsonPath = "flattened.json";

        try
        {
            // Verify that the input Excel file exists
            if (!File.Exists(inputExcelPath))
            {
                Console.WriteLine($"Error: Input file '{inputExcelPath}' not found.");
                return;
            }

            // Load the workbook using Aspose.Cells
            var workbook = new Workbook(inputExcelPath);

            // Export workbook to JSON string via a memory stream
            string jsonContent;
            using (var ms = new MemoryStream())
            {
                var saveOptions = new JsonSaveOptions(); // Correct save options for JSON
                workbook.Save(ms, saveOptions);
                ms.Position = 0;
                using var reader = new StreamReader(ms, Encoding.UTF8);
                jsonContent = reader.ReadToEnd();
            }

            // Parse the JSON document
            using JsonDocument doc = JsonDocument.Parse(jsonContent);
            JsonElement root = doc.RootElement;

            // Flatten the JSON structure
            var flatDictionary = new Dictionary<string, object?>();
            FlattenElement(root, "", flatDictionary);

            // Serialize the flattened dictionary back to JSON
            var options = new JsonSerializerOptions { WriteIndented = true };
            string flattenedJson = JsonSerializer.Serialize(flatDictionary, options);

            // Write the flattened JSON to the output file
            File.WriteAllText(outputJsonPath, flattenedJson, Encoding.UTF8);

            Console.WriteLine($"Flattened JSON has been saved to '{outputJsonPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    /// <param name="element">The current JsonElement to process.</param>
    /// <param name="prefix">The accumulated property path.</param>
    /// <param name="result">The dictionary collecting flattened key‑value pairs.</param>
    static void FlattenElement(JsonElement element, string prefix, Dictionary<string, object?> result)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                foreach (JsonProperty property in element.EnumerateObject())
                {
                    string newPrefix = string.IsNullOrEmpty(prefix) ? property.Name : $"{prefix}.{property.Name}";
                    FlattenElement(property.Value, newPrefix, result);
                }
                break;

            case JsonValueKind.Array:
                int index = 0;
                foreach (JsonElement item in element.EnumerateArray())
                {
                    string newPrefix = $"{prefix}[{index}]";
                    FlattenElement(item, newPrefix, result);
                    index++;
                }
                break;

            case JsonValueKind.String:
                result[prefix] = element.GetString();
                break;

            case JsonValueKind.Number:
                if (element.TryGetInt64(out long l))
                    result[prefix] = l;
                else if (element.TryGetDouble(out double d))
                    result[prefix] = d;
                else
                    result[prefix] = element.GetDecimal();
                break;

            case JsonValueKind.True:
            case JsonValueKind.False:
                result[prefix] = element.GetBoolean();
                break;

            case JsonValueKind.Null:
                result[prefix] = null;
                break;

            default:
                // For other kinds (e.g., Undefined), store the raw text
                result[prefix] = element.GetRawText();
                break;
        }
    }
}
