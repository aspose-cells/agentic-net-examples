// Title: Export Aspose.Cells Workbook to camelCase JSON in C#
// Description: Creates an Excel workbook, fills a sheet with employee data, uses Aspose.Cells JsonSaveOptions to export the used range as a pretty‑printed JSON object, then converts all property names to camelCase with System.Text.Json.Nodes and writes the result to a file.
// Keywords: Aspose.Cells | C# export to JSON | JsonSaveOptions | camelCase JSON | ExportRangeToJson | Excel to JSON | pretty printed JSON | System.Text.Json.Nodes | header row as property names | convert JSON keys
// Common Searches: Aspose.Cells export worksheet to JSON camelCase | C# convert exported JSON keys to camelCase | JsonSaveOptions AlwaysExportAsJsonObject example | how to use JsonUtility.ExportRangeToJson | pretty print JSON from Excel C#
// Developer Intent: Generate a JSON file from an Excel workbook where all keys follow camelCase naming.
// Use Cases: Provide API payloads for JavaScript front‑ends that require camelCase fields. | Create configuration files from Excel templates with camelCase keys for downstream services. | Integrate Excel‑based data into microservices that expect camelCase JSON. | Automate data exchange between Excel and Node.js applications. | Standardize naming conventions for data pipelines consuming Excel exports.
// AI Prompts: Write a C# method that recursively converts JSON object property names from PascalCase to camelCase using System.Text.Json.Nodes. | Show how to configure Aspose.Cells JsonSaveOptions to export a worksheet range as a pretty‑printed JSON object with the first row treated as headers. | Demonstrate robust error handling when exporting a workbook to JSON and then transforming property names to camelCase. | Provide a complete C# console program that creates a workbook, populates employee data, exports to JSON, and converts keys to camelCase.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using System.Text.Json;
using System.Text.Json.Nodes;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsJsonCamelCase
{
    // Creates an Excel workbook, fills a sheet with employee data, uses Aspose.Cells JsonSaveOptions to export the used range as a pretty‑printed JSON object, then converts all property names to camelCase with System.Text.Json.Nodes and writes the result to a file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Employees";

                // Header row (will become JSON property names)
                sheet.Cells["A1"].PutValue("EmployeeId");
                sheet.Cells["B1"].PutValue("FirstName");
                sheet.Cells["C1"].PutValue("LastName");
                sheet.Cells["D1"].PutValue("Salary");

                // Data rows
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("John");
                sheet.Cells["C2"].PutValue("Doe");
                sheet.Cells["D2"].PutValue(50000);

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Jane");
                sheet.Cells["C3"].PutValue("Smith");
                sheet.Cells["D3"].PutValue(60000);

                // Configure JSON save options
                JsonSaveOptions saveOptions = new JsonSaveOptions
                {
                    // Export as an object even if only one sheet exists
                    AlwaysExportAsJsonObject = true,
                    // Treat the first row as header (property names)
                    HasHeaderRow = true,
                    // Pretty‑print JSON with two‑space indentation
                    Indent = "  ",
                    // Export empty cells as null (optional)
                    ExportEmptyCells = true
                };

                // Export the used range of the worksheet to a JSON string
                AsposeRange exportRange = sheet.Cells.MaxDisplayRange;
                string rawJson = JsonUtility.ExportRangeToJson(exportRange, saveOptions);

                // Convert all JSON property names to camelCase
                string camelCaseJson = ConvertPropertyNamesToCamelCase(rawJson);

                // Write the camelCase JSON to a file
                string outputPath = "EmployeesCamelCase.json";
                File.WriteAllText(outputPath, camelCaseJson);

                Console.WriteLine($"JSON exported to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Parses JSON, transforms property names to camelCase, and returns the formatted string
        static string ConvertPropertyNamesToCamelCase(string json)
        {
            try
            {
                JsonNode? rootNode = JsonNode.Parse(json);
                if (rootNode is null)
                {
                    throw new InvalidOperationException("Failed to parse JSON.");
                }

                JsonNode transformed = TransformNode(rootNode);
                var options = new JsonSerializerOptions { WriteIndented = true };
                return JsonSerializer.Serialize(transformed, options);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"JSON conversion error: {ex.Message}");
                return string.Empty;
            }
        }

        // Recursively processes a JsonNode, renaming object properties to camelCase
        static JsonNode TransformNode(JsonNode node)
        {
            if (node is JsonObject obj)
            {
                var newObj = new JsonObject();
                foreach (var kvp in obj)
                {
                    if (string.IsNullOrEmpty(kvp.Key))
                    {
                        continue;
                    }

                    string camel = char.ToLowerInvariant(kvp.Key[0]) + kvp.Key.Substring(1);
                    newObj[camel] = TransformNode(kvp.Value);
                }
                return newObj;
            }
            else if (node is JsonArray arr)
            {
                var newArr = new JsonArray();
                foreach (var item in arr)
                {
                    newArr.Add(TransformNode(item));
                }
                return newArr;
            }
            else
            {
                // Primitive value (string, number, bool, null)
                return node;
            }
        }
    }
}
