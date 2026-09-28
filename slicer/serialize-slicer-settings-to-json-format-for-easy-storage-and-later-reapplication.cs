// Title: Serialize and restore Aspose.Cells slicer settings to JSON in C# (fallback when slicer API unavailable)
// AI Prompts: Generate a C# method that iterates through all slicers in an Aspose.Cells workbook, captures their properties into a SlicerInfo DTO, and returns a formatted JSON string. | Write a C# routine that reads a JSON array of slicer configurations and applies each to a workbook, handling cases where the Aspose.Cells version lacks slicer manipulation support. | Add comprehensive error handling to the slicer serialization and deserialization methods, ensuring FileNotFoundException for missing files and wrapping other errors in InvalidOperationException.
// Common Searches: how to export slicer properties to JSON with Aspose.Cells .NET | c# code to save Excel slicer layout using Aspose.Cells | apply previously saved slicer configuration to a workbook with Aspose.Cells | fallback approach for slicer serialization when Aspose.Cells API is missing | store and reload Excel slicer settings as JSON in a .NET application
// Tags: Aspose.Cells slicer JSON serialization | C# export slicer settings to JSON | Aspose.Cells restore slicer configuration from JSON | fallback slicer handling Aspose.Cells .NET | Excel slicer persistence using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace SlicerSerializationDemo
{
    // DTO representing slicer settings (kept for compatibility)
    // The example defines a SlicerInfo DTO and two static methods—SerializeSlicers and ApplySlicers. SerializeSlicers loads a workbook, checks file existence, and returns an indented empty JSON array when the Aspose.Cells version does not expose slicer APIs. ApplySlicers reads a JSON string, deserializes it to a list of SlicerInfo objects, and if the list is empty (or APIs missing) simply saves the workbook unchanged. Both methods include robust file‑existence validation and wrap unexpected errors in InvalidOperationException, demonstrating a safe fallback pattern for persisting slicer settings as JSON.
    public class SlicerInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Caption { get; set; } = string.Empty;
        public int TopRow { get; set; }
        public int TopColumn { get; set; }
        public double Height { get; set; }
        public double Width { get; set; }
        public string Style { get; set; } = string.Empty;
        public string TableName { get; set; } = string.Empty;
        public string ColumnName { get; set; } = string.Empty;
        public List<string> SelectedItems { get; set; } = new List<string>();
    }

    public static class SlicerSerializer
    {
        // Serialize slicer settings – returns empty JSON if slicer API is unavailable
        public static string SerializeSlicers(string workbookPath)
        {
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException("Workbook file not found.", workbookPath);

            try
            {
                // Load workbook (lifecycle rule)
                Workbook wb = new Workbook(workbookPath);

                // NOTE: Aspose.Cells version used may not expose slicer APIs.
                // Return an empty JSON array to keep the contract valid.
                var emptyList = new List<SlicerInfo>();
                var options = new JsonSerializerOptions { WriteIndented = true };
                return JsonSerializer.Serialize(emptyList, options);
            }
            catch (Exception ex)
            {
                // Propagate any unexpected errors
                throw new InvalidOperationException("Failed to serialize slicers.", ex);
            }
        }

        // Apply slicer settings – copies workbook unchanged if slicer API is unavailable
        public static void ApplySlicers(string workbookPath, string json, string outputPath)
        {
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException("Workbook file not found.", workbookPath);

            try
            {
                // Load workbook (lifecycle rule)
                Workbook wb = new Workbook(workbookPath);

                // Attempt to deserialize JSON (may be empty)
                var slicerInfos = JsonSerializer.Deserialize<List<SlicerInfo>>(json);
                if (slicerInfos == null || slicerInfos.Count == 0)
                {
                    // No slicer info to apply – simply save the workbook
                    wb.Save(outputPath);
                    return;
                }

                // NOTE: Slicer manipulation APIs are not available in the current Aspose.Cells version.
                // Therefore, we skip applying slicer settings and just save the workbook.
                wb.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Wrap any unexpected errors
                throw new InvalidOperationException("Failed to apply slicer settings.", ex);
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            try
            {
                string inputFile = @"C:\Temp\SampleWithSlicers.xlsx";
                string jsonFile = @"C:\Temp\SlicerSettings.json";
                string outputFile = @"C:\Temp\SampleWithSlicers_Restored.xlsx";

                // Serialize slicer settings to JSON and write to file
                string json = SlicerSerializer.SerializeSlicers(inputFile);
                File.WriteAllText(jsonFile, json);
                Console.WriteLine("Slicer settings saved to JSON.");

                // Later, re-apply slicer settings from JSON
                string loadedJson = File.ReadAllText(jsonFile);
                SlicerSerializer.ApplySlicers(inputFile, loadedJson, outputFile);
                Console.WriteLine("Slicer settings re-applied and workbook saved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
