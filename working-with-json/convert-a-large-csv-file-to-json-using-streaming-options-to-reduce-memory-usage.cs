// Title: Convert Large CSV to JSON with Aspose.Cells Streaming LoadOptions (C#)
// Description: Shows how to stream a massive CSV file into an Aspose.Cells Workbook using LoadOptions.MemoryPreference, then export it as JSON with JsonSaveOptions (header row, empty cells). The method uses FileStream to keep RAM usage low and includes basic error handling.
// Keywords: Aspose.Cells | C# | CSV to JSON | streaming load | MemoryPreference | MemorySetting | JsonSaveOptions | large file conversion | low memory | FileStream | LoadOptions | ExportEmptyCells | HasHeaderRow
// Common Searches: convert large csv to json c# without loading whole file | aspnet streaming loadoptions csv memorypreference | aspose.cells save workbook as json with header row | export empty cells as null aspose.cells jsonsaveoptions | low‑memory csv to json conversion .net
// Developer Intent: Create a JSON file from a huge CSV while minimizing memory consumption using Aspose.Cells streaming features.
// Use Cases: Process multi‑gigabyte log files into JSON for API ingestion without exhausting server RAM. | Generate front‑end data payloads from legacy CSV reports in a memory‑efficient way. | Migrate on‑premises CSV exports to cloud storage in JSON format using streaming to handle very large datasets.
// AI Prompts: Write C# code that reads a CSV with Aspose.Cells using MemoryPreference and saves it to JSON with a custom date format. | Explain the effect of MemorySetting.MemoryPreference on memory usage during large file conversion in Aspose.Cells. | Show how to configure JsonSaveOptions to omit empty cells and flatten the output when converting CSV to JSON.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Shows how to stream a massive CSV file into an Aspose.Cells Workbook using LoadOptions.MemoryPreference, then export it as JSON with JsonSaveOptions (header row, empty cells). The method uses FileStream to keep RAM usage low and includes basic error handling.
class CsvToJsonStreaming
{
    static void Main()
    {
        // Path to the large CSV file
        string csvPath = "large_input.csv";

        // Path where the resulting JSON will be saved
        string jsonPath = "output.json";

        // Verify that the CSV file exists before proceeding
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"CSV file not found: {csvPath}");
            return;
        }

        try
        {
            // Configure load options for CSV with memory optimization (streaming)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv)
            {
                MemorySetting = MemorySetting.MemoryPreference // reduces memory consumption
            };

            // Open the CSV file as a stream to avoid loading the whole file into memory
            using (FileStream csvStream = new FileStream(csvPath, FileMode.Open, FileAccess.Read))
            {
                // Load the CSV data into a workbook using the streaming load options
                Workbook workbook = new Workbook(csvStream, loadOptions);

                // Configure JSON save options (customize as needed)
                JsonSaveOptions jsonSaveOptions = new JsonSaveOptions
                {
                    ExportEmptyCells = true,   // include empty cells as null
                    HasHeaderRow = true        // treat first row as header
                };

                // Save the workbook as a JSON file
                workbook.Save(jsonPath, jsonSaveOptions);
            }

            Console.WriteLine("CSV file has been successfully converted to JSON.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
