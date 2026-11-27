// Title: Flatten a nested JSON file to dot‑separated keys and export the result to an XLSX workbook using Aspose.Cells in C#
// Common Searches: C# code to flatten hierarchical JSON into a single‑level JSON with dot notation | How to convert nested JSON to flat JSON and save as .json file using System.Text.Json | Import flattened JSON into an Excel worksheet with Aspose.Cells JsonUtility in C# | Save flattened JSON data to an XLSX file using Aspose.Cells Workbook in .NET | Command‑line tool for flattening JSON and generating an Excel report with Aspose.Cells
// Tags: json flattening with System.Text.Json in C# | Aspose.Cells JsonUtility import flattened JSON to worksheet | export flat JSON to XLSX using Aspose.Cells | dot‑separated key generation from nested JSON | command‑line JSON flattener .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace JsonFlattener
{
    // The sample reads a JSON file, recursively flattens nested objects and arrays into dot‑ and bracket‑separated keys, writes the flattened JSON to an output file, and optionally imports the flattened JSON into an Excel worksheet using Aspose.Cells' JsonUtility before saving the workbook as an XLSX file.
    class Program
    {
        /// <param name="args">
        /// args[0] – path to the source JSON file (optional, defaults to "input.json").
        /// args[1] – path to the flattened JSON output file (optional, defaults to "flattened.json").
        /// args[2] – path to the Excel workbook that will contain the flattened data (optional, defaults to "flattened.xlsx").
        /// </param>
        static void Main(string[] args)
        {
            // Resolve file paths
            string inputJsonPath = args.Length > 0 ? args[0] : "input.json";
            string outputJsonPath = args.Length > 1 ? args[1] : "flattened.json";
            string outputExcelPath = args.Length > 2 ? args[2] : "flattened.xlsx";

            // Validate input file existence
            if (!File.Exists(inputJsonPath))
            {
                Console.Error.WriteLine($"Error: Input JSON file not found at \"{Path.GetFullPath(inputJsonPath)}\".");
                return;
            }

            try
            {
                // Read the original JSON content
                string originalJson = File.ReadAllText(inputJsonPath, Encoding.UTF8);

                // Parse the JSON into a JsonNode hierarchy
                JsonNode rootNode = JsonNode.Parse(originalJson);
                if (rootNode == null)
                {
                    Console.Error.WriteLine("Error: Failed to parse the input JSON.");
                    return;
                }

                // Flatten the JSON structure
                var flatDictionary = new Dictionary<string, JsonNode>();
                FlattenJsonNode(rootNode, string.Empty, flatDictionary);

                // Build a new JsonObject from the flattened dictionary
                JsonObject flattenedObject = new JsonObject();
                foreach (var kvp in flatDictionary)
                {
                    flattenedObject[kvp.Key] = kvp.Value;
                }

                // Serialize the flattened JSON with indentation for readability
                string flattenedJson = flattenedObject.ToJsonString(new JsonSerializerOptions { WriteIndented = true });

                // Write the flattened JSON to the output file
                File.WriteAllText(outputJsonPath, flattenedJson, Encoding.UTF8);
                Console.WriteLine($"Flattened JSON saved to \"{Path.GetFullPath(outputJsonPath)}\".");

                // OPTIONAL: Import the flattened JSON into an Excel worksheet using Aspose.Cells
                // This demonstrates the use of Aspose.Cells' JsonUtility.ImportData method.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure layout options – we want a simple table without titles.
                JsonLayoutOptions layoutOptions = new JsonLayoutOptions
                {
                    ArrayAsTable = true,
                    IgnoreTitle = true,
                    ConvertNumericOrDate = true
                };

                // Import the flattened JSON string into the worksheet starting at cell A1 (row 0, column 0)
                JsonUtility.ImportData(flattenedJson, worksheet.Cells, 0, 0, layoutOptions);

                // Save the workbook
                workbook.Save(outputExcelPath, SaveFormat.Xlsx);
                Console.WriteLine($"Excel workbook with flattened data saved to \"{Path.GetFullPath(outputExcelPath)}\".");
            }
            catch (JsonException jsonEx)
            {
                Console.Error.WriteLine($"JSON processing error: {jsonEx.Message}");
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

        /// <summary>
        /// Recursively flattens a JsonNode hierarchy.
        /// </summary>
        /// <param name="node">Current JsonNode to process.</param>
        /// <param name="prefix">Current property path prefix (dot‑separated).</param>
        /// <param name="result">Dictionary that accumulates flattened key‑value pairs.</param>
        private static void FlattenJsonNode(JsonNode node, string prefix, Dictionary<string, JsonNode> result)
        {
            switch (node)
            {
                case JsonObject obj:
                    foreach (var property in obj)
                    {
                        string childPrefix = string.IsNullOrEmpty(prefix) ? property.Key : $"{prefix}.{property.Key}";
                        FlattenJsonNode(property.Value, childPrefix, result);
                    }
                    break;

                case JsonArray arr:
                    for (int i = 0; i < arr.Count; i++)
                    {
                        string childPrefix = $"{prefix}[{i}]";
                        FlattenJsonNode(arr[i], childPrefix, result);
                    }
                    break;

                default:
                    // Primitive value (string, number, bool, null)
                    result[prefix] = node;
                    break;
            }
        }
    }
}
