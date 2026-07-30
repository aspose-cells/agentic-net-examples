// Title: Apply Worksheet Protection from a JSON File using Aspose.Cells for .NET (C#)
// Description: A C# example that loads a workbook (or creates one), reads a JSON array of sheet‑protection settings, finds each worksheet by name, selects the appropriate ProtectionType (defaulting to All), applies optional passwords, sets fine‑grained flags such as AllowEditingObject or AllowEditingScenario, and saves the protected file.
// Keywords: Aspose.Cells worksheet protection C# | JSON sheet protection configuration | protect Excel sheet with password Aspose.Cells | ProtectionType enum Aspose.Cells | fine‑grained worksheet protection flags | load protection settings from JSON | C# Aspose.Cells code sample | GitHub Aspose.Cells protection example
// Common Searches: Aspose.Cells protect worksheet from JSON | C# read sheet protection settings JSON Aspose.Cells | apply password protection to specific worksheets Aspose.Cells | set AllowEditingObject flag using Aspose.Cells .NET | default ProtectionType All Aspose.Cells example
// Developer Intent: Read a JSON configuration file and programmatically apply the defined protection options to matching worksheets in an Excel workbook with Aspose.Cells.
// Use Cases: Bulk‑apply different protection types and passwords to multiple sheets based on an external JSON file. | Skip non‑existent worksheets while still processing the remaining configuration entries. | Adjust individual protection flags (e.g., AllowEditingObject, AllowEditingScenario) after invoking Worksheet.Protect. | Create a protected workbook from a template or an empty workbook when the source file is missing.
// AI Prompts: Generate C# code that reads a JSON file of sheet protection settings and applies them to a workbook with Aspose.Cells, handling missing sheets and defaulting to ProtectionType.All. | Create a sample protectionConfig.json for three worksheets, each with a distinct password, ProtectionType, and AllowEditingObject value, and show the expected protected workbook output. | Explain how to extend the SheetProtectionConfig class to include additional Aspose.Cells protection options such as AllowFormattingCell, and modify the example to apply those new settings.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    // Model representing protection settings for a worksheet
    // A C# example that loads a workbook (or creates one), reads a JSON array of sheet‑protection settings, finds each worksheet by name, selects the appropriate ProtectionType (defaulting to All), applies optional passwords, sets fine‑grained flags such as AllowEditingObject or AllowEditingScenario, and saves the protected file.
    public class SheetProtectionConfig
    {
        public string SheetName { get; set; }               // Target worksheet name
        public string Password { get; set; }                // Optional password
        public string ProtectionType { get; set; }          // e.g., "All", "Objects", "Scenarios"
        public bool? AllowEditingObject { get; set; }       // Optional flag
        public bool? AllowEditingScenario { get; set; }     // Optional flag
        // Add more properties as needed (e.g., AllowFormattingCell, etc.)
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Path to the existing workbook that will receive protection settings
                const string workbookPath = "template.xlsx";

                // Path to the JSON file containing protection configurations
                const string jsonConfigPath = "protectionConfig.json";

                // Load the workbook (create a new one if the template does not exist)
                Workbook workbook = File.Exists(workbookPath) ? new Workbook(workbookPath) : new Workbook();

                // Read JSON configuration if the file exists
                List<SheetProtectionConfig> configs = new List<SheetProtectionConfig>();
                if (File.Exists(jsonConfigPath))
                {
                    string json = File.ReadAllText(jsonConfigPath);
                    configs = JsonSerializer.Deserialize<List<SheetProtectionConfig>>(json) ?? new List<SheetProtectionConfig>();
                }
                else
                {
                    Console.WriteLine($"Configuration file \"{jsonConfigPath}\" not found. No protection will be applied.");
                }

                // Apply protection settings to each specified worksheet
                foreach (SheetProtectionConfig cfg in configs)
                {
                    if (cfg == null || string.IsNullOrWhiteSpace(cfg.SheetName))
                        continue;

                    // Find worksheet by name; skip if not found
                    Worksheet sheet = workbook.Worksheets[cfg.SheetName];
                    if (sheet == null)
                    {
                        Console.WriteLine($"Worksheet \"{cfg.SheetName}\" not found. Skipping.");
                        continue;
                    }

                    // Determine protection type; default to All if parsing fails
                    ProtectionType type = ProtectionType.All;
                    if (!string.IsNullOrWhiteSpace(cfg.ProtectionType) &&
                        Enum.TryParse(cfg.ProtectionType, out ProtectionType parsedType))
                    {
                        type = parsedType;
                    }

                    // Apply protection with or without password
                    if (!string.IsNullOrWhiteSpace(cfg.Password))
                    {
                        sheet.Protect(type, cfg.Password, null);
                    }
                    else
                    {
                        sheet.Protect(type);
                    }

                    // Set additional fine‑grained protection options if provided
                    Protection protection = sheet.Protection;
                    if (cfg.AllowEditingObject.HasValue)
                        protection.AllowEditingObject = cfg.AllowEditingObject.Value;
                    if (cfg.AllowEditingScenario.HasValue)
                        protection.AllowEditingScenario = cfg.AllowEditingScenario.Value;

                    // Additional flags can be set here similarly
                }

                // Save the workbook with applied protections
                const string outputPath = "protectedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with applied protection settings to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
