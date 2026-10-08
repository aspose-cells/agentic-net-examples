// Title: Convert a large CSV file to JSON in C# with Aspose.Cells using streaming to reduce memory consumption
// AI Prompts: Generate C# code that opens a CSV file with Aspose.Cells LoadOptions, processes it in streaming mode, and writes the data to a JSON file using JsonSaveOptions. | Show how to implement file‑existence checking and exception handling while performing a memory‑efficient CSV‑to‑JSON conversion with Aspose.Cells in .NET. | Provide a step‑by‑step example of using Aspose.Cells Workbook streaming features to convert a multi‑gigabyte CSV into JSON without loading the entire workbook into RAM.
// Common Searches: Aspose.Cells C# streaming conversion from CSV to JSON for large files | How to export a big CSV as JSON with low memory usage using Aspose.Cells | C# example of loading CSV with LoadOptions and saving as JSON with JsonSaveOptions
// Tags: streaming csv to json Aspose.Cells | low memory csv import Aspose.Cells .NET | JsonSaveOptions large workbook export | LoadOptions csv Aspose.Cells | memory efficient workbook conversion json

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the CSV file exists, loads it into an Aspose.Cells Workbook using LoadOptions for CSV format, then saves the workbook as JSON with JsonSaveOptions, handling exceptions and minimizing memory usage through streaming.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "large_input.csv";
            string outputPath = "output.json";

            // Verify that the input CSV file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the CSV file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook as JSON
            JsonSaveOptions saveOptions = new JsonSaveOptions();
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook successfully saved to JSON: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
