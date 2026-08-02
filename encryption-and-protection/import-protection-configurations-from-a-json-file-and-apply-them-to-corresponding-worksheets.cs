// Title: Import worksheet protection settings from JSON and apply them with Aspose.Cells for .NET
// Description: Loads an existing Excel workbook, reads a JSON file that defines sheet names, protection types and optional passwords, maps the strings to Aspose.Cells' ProtectionType enum, applies the protection to each matching worksheet, and saves the result as a new file.
// Keywords: Aspose.Cells protect worksheet JSON | C# import protection settings | Excel sheet protection programmatically | ProtectionType enum Aspose.Cells | apply password to worksheet .NET | bulk worksheet protection | JSON to Excel protection mapping
// Common Searches: how to protect Excel sheets using Aspose.Cells and a JSON file | Aspose.Cells read protection configuration from external file | C# protect multiple worksheets with different passwords | parse ProtectionType string for Aspose.Cells worksheet protection | sample code for applying worksheet protection from JSON
// Developer Intent: Read a JSON configuration that lists worksheet names, protection types and optional passwords, then programmatically protect the corresponding sheets in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Automate bulk protection of several worksheets with distinct passwords before distributing a workbook. | Enforce read‑only, objects‑only, or full protection on specific sheets during scheduled report generation. | Validate user‑editable JSON protection policies and apply only to existing worksheets, skipping missing ones.
// AI Prompts: Generate C# code that reads a JSON array of sheet protection settings and applies them to an Aspose.Cells workbook, handling missing sheets and invalid ProtectionType values. | Show an example protectionConfig.json that matches the SheetProtectionConfig model with various ProtectionType values and passwords. | Explain how to extend the snippet to also set AllowEditObjects or other AllowEdit flags after protecting a sheet with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsProtectionImport
{
    // Model representing a protection configuration for a worksheet
    // Loads an existing Excel workbook, reads a JSON file that defines sheet names, protection types and optional passwords, maps the strings to Aspose.Cells' ProtectionType enum, applies the protection to each matching worksheet, and saves the result as a new file.
    public class SheetProtectionConfig
    {
        public string SheetName { get; set; } = string.Empty;
        public string ProtectionType { get; set; } = string.Empty;   // e.g., "Objects", "All", etc.
        public string? Password { get; set; }                       // optional, can be null or empty
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Paths to the existing workbook and the JSON configuration file
                string workbookPath = "input.xlsx";
                string jsonConfigPath = "protectionConfig.json";

                // Verify that the workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file \"{workbookPath}\" not found.");
                    return;
                }

                // Verify that the JSON configuration file exists
                if (!File.Exists(jsonConfigPath))
                {
                    Console.WriteLine($"Configuration file \"{jsonConfigPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Read and deserialize JSON content
                string json = File.ReadAllText(jsonConfigPath);
                List<SheetProtectionConfig>? configs = JsonSerializer.Deserialize<List<SheetProtectionConfig>>(json);

                if (configs == null || configs.Count == 0)
                {
                    Console.WriteLine("No protection configurations found in the JSON file.");
                    return;
                }

                // Apply each protection setting to the corresponding worksheet
                foreach (var cfg in configs)
                {
                    try
                    {
                        // Find worksheet by name; if not found, skip
                        Worksheet? sheet = workbook.Worksheets[cfg.SheetName];
                        if (sheet == null)
                        {
                            Console.WriteLine($"Worksheet \"{cfg.SheetName}\" not found. Skipping.");
                            continue;
                        }

                        // Convert string to ProtectionType enum; default to All if parsing fails
                        if (!Enum.TryParse(cfg.ProtectionType, true, out ProtectionType type))
                        {
                            Console.WriteLine($"Invalid ProtectionType \"{cfg.ProtectionType}\" for sheet \"{cfg.SheetName}\". Using All.");
                            type = ProtectionType.All;
                        }

                        // Apply protection using the appropriate overload
                        if (string.IsNullOrEmpty(cfg.Password))
                        {
                            sheet.Protect(type);
                        }
                        else
                        {
                            // Aspose.Cells requires oldPassword parameter; pass empty string when not changing it
                            sheet.Protect(type, cfg.Password, string.Empty);
                        }

                        Console.WriteLine($"Applied protection to sheet \"{cfg.SheetName}\" with type {type}.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to protect sheet \"{cfg.SheetName}\": {ex.Message}");
                    }
                }

                // Save the modified workbook
                workbook.Save("output.xlsx");
                Console.WriteLine("Workbook saved as \"output.xlsx\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
