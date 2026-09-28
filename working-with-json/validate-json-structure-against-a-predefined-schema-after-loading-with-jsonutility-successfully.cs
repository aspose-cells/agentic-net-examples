// Title: Validate a JSON file against a custom property schema in C# and write the raw JSON to an Aspose.Cells worksheet
// AI Prompts: Generate C# code that reads a JSON file, validates required and optional properties using a dictionary of expected .NET types, stores the raw JSON string in cell A1 of a new Aspose.Cells worksheet, and saves the workbook as an XLSX file. | Create a C# routine that extends the validation to array‑type properties, captures any type mismatches, and writes detailed error messages to a separate worksheet in the same Aspose.Cells workbook. | Refactor the manual type‑checking logic to use System.Text.Json's built‑in schema validation features and record the overall validation outcome in a second sheet of the Excel file.
// Common Searches: c# read json file and validate required fields using a dictionary of .NET types | aspocells write json string to cell a1 and save workbook as xlsx | how to check json property types with System.Text.Json in .NET 6 without third‑party libraries | validate json against custom schema and export validation results to Excel using Aspose.Cells | c# log json validation errors into a separate worksheet in an Excel file
// Tags: c# json property validation with System.Text.Json | aspocells write raw json to worksheet | excel workbook save json validation result | custom json schema dictionary validation c# | type checking json elements .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example reads a JSON file, places its raw content into cell A1 of a new Aspose.Cells workbook, defines required and optional property dictionaries, parses the JSON with System.Text.Json, validates each property's presence and .NET type via a helper method, reports any validation errors to the console, and finally saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the JSON source file exists
            const string jsonPath = "data.json";
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"Error: JSON file '{jsonPath}' not found.");
                return;
            }

            // Load JSON content from the file
            string jsonContent = File.ReadAllText(jsonPath);

            // Create a new workbook and import JSON data into the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            try
            {
                // Aspose.Cells does not have JsonUtility in this version; write JSON string to cell A1 as fallback
                sheet.Cells[0, 0].PutValue(jsonContent);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error importing JSON into workbook: {ex.Message}");
                return;
            }

            // Simple JSON schema definition (used for manual validation)
            var requiredProperties = new Dictionary<string, Type>
            {
                { "Name", typeof(string) },
                { "Age", typeof(int) }
            };
            var optionalProperties = new Dictionary<string, Type>
            {
                { "Email", typeof(string) }
            };

            // Parse the JSON content
            using JsonDocument doc = JsonDocument.Parse(jsonContent);
            JsonElement root = doc.RootElement;

            // Validate required properties and their types
            List<string> validationErrors = new List<string>();
            foreach (var kvp in requiredProperties)
            {
                if (!root.TryGetProperty(kvp.Key, out JsonElement prop))
                {
                    validationErrors.Add($"Missing required property: {kvp.Key}");
                    continue;
                }

                if (!IsJsonElementOfType(prop, kvp.Value))
                {
                    validationErrors.Add($"Property '{kvp.Key}' is not of expected type {kvp.Value.Name}");
                }
            }

            // Validate optional properties if they exist
            foreach (var kvp in optionalProperties)
            {
                if (root.TryGetProperty(kvp.Key, out JsonElement prop) &&
                    !IsJsonElementOfType(prop, kvp.Value))
                {
                    validationErrors.Add($"Property '{kvp.Key}' is not of expected type {kvp.Value.Name}");
                }
            }

            // Output validation result
            if (validationErrors.Count == 0)
            {
                Console.WriteLine("JSON is valid according to the simple schema.");
            }
            else
            {
                Console.WriteLine("JSON validation failed. Errors:");
                foreach (string error in validationErrors)
                {
                    Console.WriteLine("- " + error);
                }
            }

            // Save the workbook
            const string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }

    // Helper method to map JsonElement kinds to .NET types
    private static bool IsJsonElementOfType(JsonElement element, Type targetType)
    {
        return targetType == typeof(string) && element.ValueKind == JsonValueKind.String ||
               targetType == typeof(int) && element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out _);
    }
}
