// Title: Import worksheet protection settings from a JSON file and apply them to an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Read a JSON array of ProtectionConfig objects and protect each matching worksheet in an Aspose.Cells Workbook, using the specified password and protection options. | Extend the code to also set AllowEditRanges and AllowDeleteRows on worksheets based on additional fields in the JSON configuration. | Add detailed logging that records which worksheets were protected, which were unprotected, and any missing sheets while processing the JSON file.
// Common Searches: how to use Aspose.Cells to protect Excel worksheets based on a JSON configuration in C# | apply password and edit permissions to specific sheets in an Excel file using Aspose.Cells .NET | load worksheet protection options from external JSON and set AllowEditObjects in Aspose.Cells | unprotect Excel worksheets programmatically with Aspose.Cells when a flag is false in config file | C# example for batch protecting multiple worksheets using Aspose.Cells and JSON settings
// Tags: batch worksheet protection Aspose.Cells JSON | set worksheet password Aspose.Cells .NET | configure AllowEditObjects Aspose.Cells | unprotect Excel sheets via Aspose.Cells config | load protection settings from JSON C#

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// Model representing protection settings for a worksheet
// The example loads an existing Excel workbook, reads a JSON file containing a list of ProtectionConfig objects, and iterates through each entry. For each matching worksheet it either protects the sheet with an optional password and attempts to set AllowEditObjects and AllowEditScenarios via reflection, or removes protection when the Protect flag is false. The modified workbook is saved to a new file, with error handling for missing files, invalid configurations, and save failures.
public class ProtectionConfig
{
    public string SheetName { get; set; } = string.Empty;   // Target worksheet name
    public bool Protect { get; set; }                       // Whether to protect the sheet
    public string? Password { get; set; }                  // Password for protection (optional)
    public bool AllowEditObjects { get; set; }              // Allow editing objects
    public bool AllowEditScenarios { get; set; }            // Allow editing scenarios
    // Additional protection options can be added as needed
}

class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string jsonPath = "protectionConfig.json";
            const string outputPath = "output.xlsx";

            // Verify required files exist
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input workbook '{inputPath}' not found.");
                return;
            }
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"Configuration file '{jsonPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Read protection configurations from JSON file
            string json = File.ReadAllText(jsonPath);
            List<ProtectionConfig>? configs = JsonSerializer.Deserialize<List<ProtectionConfig>>(json);
            if (configs == null)
            {
                Console.WriteLine("No protection configurations found.");
                return;
            }

            // Apply each configuration to its corresponding worksheet
            foreach (ProtectionConfig cfg in configs)
            {
                if (string.IsNullOrEmpty(cfg.SheetName))
                    continue; // Skip invalid entries

                Worksheet? sheet = workbook.Worksheets[cfg.SheetName];
                if (sheet == null)
                    continue; // Skip if worksheet not found

                if (cfg.Protect)
                {
                    // Protect the worksheet with optional password (oldPassword not required)
                    if (!string.IsNullOrEmpty(cfg.Password))
                        sheet.Protect(ProtectionType.All, cfg.Password, string.Empty);
                    else
                        sheet.Protect(ProtectionType.All, string.Empty, string.Empty);

                    // Attempt to set additional protection options via reflection (if supported)
                    try
                    {
                        var protection = sheet.Protection;
                        var propObj = protection.GetType().GetProperty("AllowEditObject");
                        if (propObj != null && propObj.CanWrite)
                            propObj.SetValue(protection, cfg.AllowEditObjects);

                        var propScen = protection.GetType().GetProperty("AllowEditScenario");
                        if (propScen != null && propScen.CanWrite)
                            propScen.SetValue(protection, cfg.AllowEditScenarios);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Unable to set detailed options for sheet '{cfg.SheetName}'. {ex.Message}");
                    }
                }
                else
                {
                    // Remove protection if Protect flag is false
                    sheet.Unprotect();
                }
            }

            // Save the modified workbook
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
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
