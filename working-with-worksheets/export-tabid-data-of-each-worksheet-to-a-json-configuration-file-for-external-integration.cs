// Title: Export each worksheet's name and TabId (index) to a JSON configuration file using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that loads an Excel workbook with Aspose.Cells, extracts each worksheet's Name and Index, and saves the data as an indented JSON file using System.Text.Json. | Generate code that iterates through Workbook.Worksheets, builds a list of objects containing worksheet Name and TabId, and writes the list to a JSON configuration file for external integration. | Modify the example to include additional worksheet properties (e.g., visibility, tab color) while exporting the metadata to JSON with Aspose.Cells and System.Text.Json.
// Common Searches: how to get worksheet index with Aspose.Cells and export to json in C# | c# Aspose.Cells serialize worksheet names and tab ids to json file | export Excel sheet metadata to json configuration using Aspose.Cells .NET | write worksheet name and index to json for external integration in .NET
// Tags: Aspose.Cells generate JSON configuration for worksheets | C# serialize worksheet index and name | Workbook.Worksheets JSON export | System.Text.Json Excel sheet config output | Extract worksheet TabId with Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace ExportWorksheetTabIds
{
    // Model for JSON serialization
    // The sample loads an Excel workbook with Aspose.Cells, collects each worksheet's Name and its Index (used as TabId), and writes the collection to a formatted JSON configuration file using System.Text.Json.
    public class WorksheetInfo
    {
        public string Name { get; set; }
        public int TabId { get; set; }   // Using worksheet index as TabId
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Load the Excel workbook (replace with your file path)
            string excelPath = "input.xlsx";
            Workbook workbook = new Workbook(excelPath);

            // Collect TabId information for each worksheet
            List<WorksheetInfo> sheetsInfo = new List<WorksheetInfo>();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheetsInfo.Add(new WorksheetInfo
                {
                    Name = sheet.Name,
                    TabId = sheet.Index   // Index serves as the TabId
                });
            }

            // Serialize the collection to JSON
            string json = JsonSerializer.Serialize(sheetsInfo, new JsonSerializerOptions { WriteIndented = true });

            // Save JSON to a configuration file
            string jsonPath = "worksheets_config.json";
            File.WriteAllText(jsonPath, json);

            Console.WriteLine($"Worksheet TabId data exported to '{jsonPath}'.");
        }
    }
}
