// Title: How to serialize Excel slicer settings to JSON and restore them using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export slicer selections to a JSON file | load slicer configuration from JSON and apply to an Excel workbook with Aspose.Cells | save slicer caption and selected items as JSON using the Aspose.Cells API | deserialize slicer cache items from JSON and update slicer state in .NET | persist Excel slicer settings to JSON for later reuse with Aspose.Cells
// Tags: Aspose.Cells serialize slicer settings to JSON | C# export slicer cache items JSON | Aspose.Cells deserialize slicer configuration | apply slicer selections from JSON .NET | Excel slicer state persistence Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Slicer related classes

namespace SlicerSerializationDemo
{
    // POCO classes that represent slicer settings for JSON serialization
    // The example loads an existing workbook, extracts each slicer's name, caption, and selected cache items, serializes this information into an indented JSON file, then reads the JSON back, deserializes it, finds the matching slicers by name, restores their captions and item selection states, and finally saves the workbook with the re‑applied slicer settings.
    public class SlicerItemSetting
    {
        public string Name { get; set; } = string.Empty;
        public bool Selected { get; set; }
    }

    public class SlicerSetting
    {
        public string Name { get; set; } = string.Empty;      // Slicer name (unique identifier)
        public string Caption { get; set; } = string.Empty;   // Display caption
        public List<SlicerItemSetting> Items { get; set; } = new List<SlicerItemSetting>();
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // ---------- Load existing workbook ----------
                string workbookPath = "InputWorkbook.xlsx";

                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                Workbook workbook = new Workbook(workbookPath);

                // ---------- Serialize slicer settings ----------
                List<SlicerSetting> slicerSettings = new List<SlicerSetting>();

                // Iterate through all worksheets and collect slicer information
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Slicer slicer in sheet.Slicers)
                    {
                        var setting = new SlicerSetting
                        {
                            Name = slicer.Name,
                            Caption = slicer.Caption
                        };

                        // Retrieve item states from the slicer cache
                        SlicerCache cache = slicer.SlicerCache;
                        foreach (SlicerCacheItem cacheItem in cache.SlicerCacheItems)
                        {
                            // Use the cache item's displayed value as its identifier
                            string itemName = cacheItem.Value?.ToString() ?? string.Empty;

                            setting.Items.Add(new SlicerItemSetting
                            {
                                Name = itemName,
                                Selected = cacheItem.Selected
                            });
                        }

                        slicerSettings.Add(setting);
                    }
                }

                // Convert the collected settings to JSON
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(slicerSettings, jsonOptions);

                // Save JSON to a file
                string jsonPath = "SlicerSettings.json";
                File.WriteAllText(jsonPath, json);
                Console.WriteLine($"Slicer settings serialized to {jsonPath}");

                // ---------- Re‑apply slicer settings from JSON ----------
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"JSON file not found: {jsonPath}");
                    return;
                }

                string loadedJson = File.ReadAllText(jsonPath);
                List<SlicerSetting>? loadedSettings = JsonSerializer.Deserialize<List<SlicerSetting>>(loadedJson, jsonOptions);

                if (loadedSettings == null)
                {
                    Console.WriteLine("Failed to deserialize slicer settings.");
                    return;
                }

                // Apply settings to the workbook
                foreach (SlicerSetting setting in loadedSettings)
                {
                    // Find the slicer by name across all worksheets
                    Slicer? targetSlicer = null;
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        targetSlicer = sheet.Slicers.FirstOrDefault(s => s.Name == setting.Name);
                        if (targetSlicer != null)
                            break;
                    }

                    if (targetSlicer == null)
                    {
                        Console.WriteLine($"Slicer '{setting.Name}' not found in workbook.");
                        continue;
                    }

                    // Apply caption
                    targetSlicer.Caption = setting.Caption;

                    // Apply item selection states
                    SlicerCache cache = targetSlicer.SlicerCache;
                    foreach (SlicerItemSetting itemSetting in setting.Items)
                    {
                        // Find the cache item whose displayed value matches the stored name
                        SlicerCacheItem? cacheItem = cache.SlicerCacheItems
                            .FirstOrDefault(ci => (ci.Value?.ToString() ?? string.Empty) == itemSetting.Name);

                        if (cacheItem != null)
                        {
                            cacheItem.Selected = itemSetting.Selected;
                        }
                    }
                }

                // Save the workbook after re‑applying slicer settings
                string outputPath = "Workbook_With_SlicerSettings.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with re‑applied slicer settings to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
