// Title: Flatten a nested JSON file to a single‑level JSON using System.Text.Json in C#
// AI Prompts: Create a C# console app that reads a JSON file, recursively flattens objects and arrays into dot‑separated keys, and writes the flat JSON to a new file. | Modify the flattening routine to generate snake_case keys instead of dot notation while keeping array indices intact. | Add a command‑line switch that omits entries with null values from the flattened JSON output.
// Common Searches: how to flatten nested JSON to dot‑separated keys with System.Text.Json in .NET | C# convert hierarchical JSON into a flat key/value dictionary | flatten JSON arrays with index notation using System.Text.Json | save flattened JSON output to a file in a C# console application | exclude null values while flattening JSON in C#
// Tags: flatten JSON with System.Text.Json C# | hierarchical key notation for JSON flattening | recursive JSON element processing C# | preserve numeric types during JSON flattening | write flattened JSON to file .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The example reads an input JSON file, parses it with System.Text.Json, recursively flattens objects and arrays into a dictionary using dot‑separated paths (including array indices), serializes the dictionary as indented JSON, and writes the result to an output file while handling parsing, I/O, and unexpected errors.
class Program
{
    static void Main()
    {
        // Paths to the source and output JSON files
        string inputPath = "input.json";
        string outputPath = "flattened.json";

        // Verify that the input file exists before attempting to read it
        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the JSON document from file
            string jsonContent = File.ReadAllText(inputPath);
            using JsonDocument doc = JsonDocument.Parse(jsonContent);

            // Flatten the JSON structure
            var flatDictionary = new Dictionary<string, object>();
            FlattenElement(doc.RootElement, "", flatDictionary);

            // Convert the flat dictionary to a formatted JSON string
            string flattenedJson = JsonSerializer.Serialize(
                flatDictionary,
                new JsonSerializerOptions { WriteIndented = true });

            // Save the flattened JSON to the output file
            File.WriteAllText(outputPath, flattenedJson);

            Console.WriteLine($"Flattened JSON has been saved to '{outputPath}'.");
        }
        catch (JsonException jsonEx)
        {
            Console.Error.WriteLine($"JSON parsing error: {jsonEx.Message}");
        }
        catch (IOException ioEx)
        {
            Console.Error.WriteLine($"I/O error: {ioEx.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    static void FlattenElement(JsonElement element, string prefix, Dictionary<string, object> result)
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
                result[prefix] = element.GetString()!;
                break;

            case JsonValueKind.Number:
                // Preserve the numeric type as best as possible
                if (element.TryGetInt64(out long l))
                    result[prefix] = l;
                else if (element.TryGetDouble(out double d))
                    result[prefix] = d;
                else
                    result[prefix] = element.GetRawText();
                break;

            case JsonValueKind.True:
            case JsonValueKind.False:
                result[prefix] = element.GetBoolean();
                break;

            case JsonValueKind.Null:
                result[prefix] = null!;
                break;

            default:
                // For any other kinds, store the raw JSON text
                result[prefix] = element.GetRawText();
                break;
        }
    }
}
