// Title: Export worksheet protection settings to a JSON file with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, reads each worksheet's Protection properties, and writes the collected flags to an indented JSON file. | Create a reusable method that uses reflection to safely obtain all boolean protection options from a Worksheet object and returns a serializable model for JSON export. | Modify the export routine to also include the worksheet password hash (if any) alongside the protection flags in the generated JSON report.
// Common Searches: Aspose.Cells C# export worksheet protection flags to JSON | how to read Excel sheet protection settings with Aspose.Cells | serialize Excel worksheet security options to a JSON file in .NET | audit protected worksheets by extracting protection options using Aspose.Cells
// Tags: Aspose.Cells worksheet protection JSON export | C# serialize Excel sheet security settings | reflection based extraction of protection flags | audit Excel worksheet protection with JSON | export workbook protection configuration .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace WorksheetProtectionExport
{
    // Model representing protection settings of a worksheet
    // The program loads an Excel workbook via Aspose.Cells, iterates through each worksheet, captures all protection-related boolean flags using reflection, and writes the aggregated data as a formatted JSON file for external auditing.
    public class WorksheetProtectionInfo
    {
        public string SheetName { get; set; } = string.Empty;
        public bool IsProtected { get; set; }
        public bool AllowDeleteColumns { get; set; }
        public bool AllowDeleteRows { get; set; }
        public bool AllowEditObject { get; set; }
        public bool AllowEditScenario { get; set; }
        public bool AllowFilter { get; set; }
        public bool AllowFormatCells { get; set; }
        public bool AllowFormatColumns { get; set; }
        public bool AllowFormatRows { get; set; }
        public bool AllowInsertColumns { get; set; }
        public bool AllowInsertHyperlinks { get; set; }
        public bool AllowInsertRows { get; set; }
        public bool AllowPivotTables { get; set; }
        public bool AllowSelectLockedCells { get; set; }
        public bool AllowSelectUnlockedCells { get; set; }
        public bool AllowSort { get; set; }
    }

    class Program
    {
        // Helper to safely read a boolean property via reflection
        private static bool GetBoolProperty(object obj, string propertyName)
        {
            try
            {
                var prop = obj.GetType().GetProperty(propertyName);
                if (prop != null && prop.PropertyType == typeof(bool))
                {
                    return (bool)prop.GetValue(obj)!;
                }
            }
            catch
            {
                // Ignore any reflection errors and fall back to false
            }
            return false;
        }

        static void Main(string[] args)
        {
            // Path to the source Excel file
            string sourceFilePath = "input.xlsx";

            // Path to the output JSON file
            string outputJsonPath = "worksheet_protection.json";

            // Verify that the source file exists
            if (!File.Exists(sourceFilePath))
            {
                Console.WriteLine($"Source file '{sourceFilePath}' not found.");
                return;
            }

            try
            {
                // Load the workbook (using Aspose.Cells)
                Workbook workbook = new Workbook(sourceFilePath);

                // List to hold protection info for each worksheet
                List<WorksheetProtectionInfo> protectionInfoList = new List<WorksheetProtectionInfo>();

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the protection object of the worksheet
                    var protection = sheet.Protection;

                    // Populate the model using reflection to avoid version‑specific API issues
                    WorksheetProtectionInfo info = new WorksheetProtectionInfo
                    {
                        SheetName = sheet.Name,
                        IsProtected = GetBoolProperty(protection, "IsProtected"),
                        AllowDeleteColumns = GetBoolProperty(protection, "AllowDeleteColumns"),
                        AllowDeleteRows = GetBoolProperty(protection, "AllowDeleteRows"),
                        AllowEditObject = GetBoolProperty(protection, "AllowEditObject"),
                        AllowEditScenario = GetBoolProperty(protection, "AllowEditScenario"),
                        AllowFilter = GetBoolProperty(protection, "AllowFilter"),
                        AllowFormatCells = GetBoolProperty(protection, "AllowFormatCells"),
                        AllowFormatColumns = GetBoolProperty(protection, "AllowFormatColumns"),
                        AllowFormatRows = GetBoolProperty(protection, "AllowFormatRows"),
                        AllowInsertColumns = GetBoolProperty(protection, "AllowInsertColumns"),
                        AllowInsertHyperlinks = GetBoolProperty(protection, "AllowInsertHyperlinks"),
                        AllowInsertRows = GetBoolProperty(protection, "AllowInsertRows"),
                        AllowPivotTables = GetBoolProperty(protection, "AllowPivotTables"),
                        AllowSelectLockedCells = GetBoolProperty(protection, "AllowSelectLockedCells"),
                        AllowSelectUnlockedCells = GetBoolProperty(protection, "AllowSelectUnlockedCells"),
                        AllowSort = GetBoolProperty(protection, "AllowSort")
                    };

                    protectionInfoList.Add(info);
                }

                // Serialize the list to JSON with indentation for readability
                JsonSerializerOptions jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                string jsonString = JsonSerializer.Serialize(protectionInfoList, jsonOptions);

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(outputJsonPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write JSON to the output file
                File.WriteAllText(outputJsonPath, jsonString);

                Console.WriteLine($"Protection settings exported to '{outputJsonPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
