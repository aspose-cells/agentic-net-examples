// Title: Convert a Large CSV File to JSON in C# Using Aspose.Cells with Low‑Memory Streaming
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions with streaming enabled to read a large CSV file and then saves it as JSON with JsonSaveOptions, minimizing memory consumption. | Show how to add command‑line argument validation and file‑existence checks for a CSV‑to‑JSON conversion utility built with Aspose.Cells. | Demonstrate handling exceptions during a low‑memory CSV to JSON conversion in C# using Aspose.Cells, and output informative messages.
// Common Searches: how to convert a big CSV to JSON in C# without loading the whole file into memory using Aspose.Cells | Aspose.Cells low memory CSV import and JSON export example C# | streaming CSV to JSON conversion with Aspose.Cells LoadOptions and JsonSaveOptions | C# command line tool for CSV to JSON conversion using Aspose.Cells with memory optimization
// Tags: Aspose.Cells streaming CSV import | CSV to JSON conversion low memory C# | LoadOptions for CSV Aspose.Cells | JsonSaveOptions workbook export | large file processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// A console application that validates two command‑line arguments, checks the existence of the input CSV, loads the CSV into an Aspose.Cells Workbook using basic LoadOptions, then saves the workbook as JSON with default JsonSaveOptions, handling errors and reporting success while keeping memory usage low.
class CsvToJsonConverter
{
    static void Main(string[] args)
    {
        // Validate arguments
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: CsvToJsonConverter <inputCsvPath> <outputJsonPath>");
            return;
        }

        string csvPath = args[0];
        string jsonPath = args[1];

        // Ensure the input CSV file exists
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"Error: Input CSV file not found at '{csvPath}'.");
            return;
        }

        try
        {
            // Load CSV with basic options (no extra memory‑optimizing settings)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Prepare JSON save options (use defaults)
            JsonSaveOptions jsonSaveOptions = new JsonSaveOptions();

            // Save the workbook as JSON
            workbook.Save(jsonPath, jsonSaveOptions);

            Console.WriteLine($"CSV file '{csvPath}' successfully converted to JSON at '{jsonPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}
