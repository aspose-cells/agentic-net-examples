// Title: How to export each worksheet's TabId to a JSON configuration file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet TabId values to JSON file | Save Excel sheet TabId and name as JSON using Aspose.Cells | Generate JSON mapping of worksheet names and TabId with Aspose.Cells .NET | Command‑line tool to write worksheet TabId list to JSON in C# | Aspose.Cells retrieve TabId property for all worksheets and serialize to JSON
// Tags: Aspose.Cells export worksheet TabId to JSON | C# serialize Excel sheet identifiers | Workbook.Worksheets TabId property usage | JSON configuration file for Excel worksheets | command line Aspose.Cells workbook processing

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace ExportWorksheetTabIds
{
    // A C# console application that loads an Excel workbook via Aspose.Cells, iterates through all worksheets to collect each sheet's Name and TabId, serializes the collection into an indented JSON file named worksheet_tabids.json placed alongside the source workbook, and reports success or errors to the console.
    class Program
    {
        public class WorksheetTabInfo
        {
            public string WorksheetName { get; set; }
            public int TabId { get; set; }
        }

        static void Main(string[] args)
        {
            // Expect exactly one argument: the path to the source Excel workbook.
            if (args.Length != 1)
            {
                Console.WriteLine("Usage: ExportWorksheetTabIds <input-workbook-path>");
                return;
            }

            string inputPath = args[0];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' does not exist.");
                return;
            }

            // Determine output JSON file path (same folder, fixed name)
            string outputPath = Path.Combine(Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory(),
                                             "worksheet_tabids.json");

            try
            {
                // Load the workbook. Using a using block ensures proper disposal.
                using (Workbook workbook = new Workbook(inputPath))
                {
                    // Collect TabId data for each worksheet.
                    List<WorksheetTabInfo> tabInfoList = new List<WorksheetTabInfo>();

                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // The TabId property is an integer identifier for the sheet.
                        tabInfoList.Add(new WorksheetTabInfo
                        {
                            WorksheetName = sheet.Name,
                            TabId = sheet.TabId
                        });
                    }

                    // Serialize the list to JSON with indentation for readability.
                    JsonSerializerOptions jsonOptions = new JsonSerializerOptions
                    {
                        WriteIndented = true
                    };

                    string jsonContent = JsonSerializer.Serialize(tabInfoList, jsonOptions);

                    // Write JSON to the output file.
                    File.WriteAllText(outputPath, jsonContent);

                    Console.WriteLine($"Successfully exported TabId data for {tabInfoList.Count} worksheet(s).");
                    Console.WriteLine($"JSON configuration saved to: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                // Do not hide failures – report the full exception message and stack trace.
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
