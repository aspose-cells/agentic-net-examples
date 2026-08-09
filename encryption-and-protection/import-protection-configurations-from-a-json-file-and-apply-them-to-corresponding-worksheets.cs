// Title: C# – Apply Worksheet Protection from JSON with Aspose.Cells
// Description: Demonstrates how to read a JSON file that defines sheet names, passwords, and protection types, locate or create the corresponding worksheets in an Aspose.Cells workbook, apply the specified protection, and save the result as an XLSX file.
// Keywords: Aspose.Cells protect worksheet C# | JSON worksheet protection Aspose | apply protection type enum Aspose.Cells | load protection settings from JSON | C# Excel sheet security Aspose | dynamic worksheet creation Aspose.Cells | protect multiple sheets programmatically
// Common Searches: Aspose.Cells protect worksheets using JSON | C# read JSON and apply Excel sheet protection | set protection type for Excel sheets Aspose | create missing worksheets and protect them C# | load protection configuration file Aspose.Cells
// Developer Intent: Read a JSON configuration and programmatically protect the matching worksheets in an Aspose.Cells workbook.
// Use Cases: Bulk‑apply passwords and protection levels to many sheets based on a JSON manifest. | Automatically generate a new worksheet when the specified name is absent and enforce the defined security settings. | Fallback to full protection when the JSON entry lacks a valid ProtectionType value.
// AI Prompts: Write C# code that loads a JSON file with worksheet protection rules and applies them using Aspose.Cells, handling missing sheets and invalid enum values. | Show how to modify the sample to open an existing workbook instead of creating a new one while still using JSON‑driven protection settings. | Provide robust error‑handling suggestions for JSON deserialization and ProtectionType parsing in the Aspose.Cells protection workflow.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Model representing each protection rule from JSON
    // Demonstrates how to read a JSON file that defines sheet names, passwords, and protection types, locate or create the corresponding worksheets in an Aspose.Cells workbook, apply the specified protection, and save the result as an XLSX file.
    public class ProtectionConfig
    {
        public string? SheetName { get; set; }
        public string? Password { get; set; }
        public string? Type { get; set; }
    }

    public static class ApplyProtectionFromJson
    {
        public static void Run()
        {
            const string jsonFilePath = "protectionConfig.json";

            // Verify that the JSON configuration file exists
            if (!File.Exists(jsonFilePath))
            {
                Console.WriteLine($"Configuration file not found: {jsonFilePath}");
                return;
            }

            try
            {
                // Read and deserialize the JSON content
                string jsonContent = File.ReadAllText(jsonFilePath);
                List<ProtectionConfig>? configs = JsonSerializer.Deserialize<List<ProtectionConfig>>(jsonContent);

                if (configs == null || configs.Count == 0)
                {
                    Console.WriteLine("No protection configurations found in the JSON file.");
                    return;
                }

                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook();

                // Ensure at least one worksheet exists
                if (workbook.Worksheets.Count == 0)
                {
                    workbook.Worksheets.Add();
                }

                // Apply each protection rule
                foreach (ProtectionConfig cfg in configs)
                {
                    if (string.IsNullOrWhiteSpace(cfg.SheetName))
                    {
                        Console.WriteLine("Skipping entry with missing SheetName.");
                        continue;
                    }

                    // Find existing worksheet or create a new one
                    Worksheet? sheet = null;
                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        if (ws.Name.Equals(cfg.SheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            sheet = ws;
                            break;
                        }
                    }

                    if (sheet == null)
                    {
                        int newIndex = workbook.Worksheets.Add();
                        sheet = workbook.Worksheets[newIndex];
                        sheet.Name = cfg.SheetName;
                    }

                    // Determine the protection type (default to All)
                    ProtectionType protectionType = ProtectionType.All;
                    if (!string.IsNullOrWhiteSpace(cfg.Type))
                    {
                        if (!Enum.TryParse<ProtectionType>(cfg.Type, true, out protectionType))
                        {
                            protectionType = ProtectionType.All;
                        }
                    }

                    // Apply protection to the worksheet
                    sheet.Protect(protectionType, cfg.Password ?? string.Empty, null);
                }

                // Save the protected workbook
                const string outputPath = "ProtectedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ApplyProtectionFromJson.Run();
        }
    }
}
