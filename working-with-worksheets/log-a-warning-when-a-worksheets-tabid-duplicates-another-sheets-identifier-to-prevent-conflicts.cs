// Title: Log a warning for duplicate worksheet TabId values in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to check for duplicate worksheet TabId in Aspose.Cells C# | C# code to log warnings when Excel sheet TabId is duplicated with Aspose.Cells | detect worksheet TabId conflicts in a .NET workbook using Aspose.Cells | Aspose.Cells example for identifying duplicate sheet TabId values | validate uniqueness of worksheet TabId before saving workbook in C#
// Tags: duplicate worksheet TabId detection Aspose.Cells | log warning for sheet identifier conflict C# | Aspose.Cells TabId uniqueness validation | C# workbook worksheet TabId mapping | Aspose.Cells worksheet TabId duplicate handling

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorksheetTabIdValidator
{
    // The sample loads an Excel workbook with Aspose.Cells, builds a dictionary of each worksheet's TabId to its names, detects any TabId values that appear on multiple sheets, writes warning messages listing the affected worksheets, and optionally saves the unchanged workbook to a specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Error: No input workbook path provided.");
                Console.WriteLine("Usage: WorksheetTabIdValidator.exe <inputWorkbookPath> [outputWorkbookPath]");
                return;
            }

            string inputPath = args[0];
            string outputPath = args.Length > 1 ? args[1] : null;

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                WorksheetCollection sheets = workbook.Worksheets;

                // Dictionary to map TabId -> list of worksheet names that share this TabId
                Dictionary<int, List<string>> tabIdMap = new Dictionary<int, List<string>>();

                // Populate the map
                foreach (Worksheet sheet in sheets)
                {
                    int tabId = sheet.TabId;
                    if (!tabIdMap.ContainsKey(tabId))
                    {
                        tabIdMap[tabId] = new List<string>();
                    }
                    tabIdMap[tabId].Add(sheet.Name);
                }

                // Detect duplicates and log warnings
                bool duplicatesFound = false;
                foreach (var kvp in tabIdMap)
                {
                    if (kvp.Value.Count > 1)
                    {
                        duplicatesFound = true;
                        Console.WriteLine($"Warning: Duplicate TabId detected: {kvp.Key}");
                        Console.WriteLine("  Affected worksheets:");
                        foreach (string sheetName in kvp.Value)
                        {
                            Console.WriteLine($"    - {sheetName}");
                        }
                    }
                }

                if (!duplicatesFound)
                {
                    Console.WriteLine("No duplicate TabId values were found.");
                }

                // Optionally save the workbook (unchanged) if an output path is supplied
                if (!string.IsNullOrWhiteSpace(outputPath))
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An exception occurred while processing the workbook: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
