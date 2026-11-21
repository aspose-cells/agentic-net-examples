// Title: Convert a large CSV file to JSON in C# with Aspose.Cells using LoadOptions for memory‑efficient processing
// Common Searches: C# Aspose.Cells load large CSV with LoadOptions and export to JSON | How to convert big CSV to JSON using Aspose.Cells without high memory usage | Aspose.Cells Workbook.Save JSON format example in .NET | Reading CSV as workbook in Aspose.Cells with LoadFormat.Csv | Convert CSV to JSON programmatically with Aspose.Cells in C#
// Tags: Aspose.Cells LoadOptions CSV processing | Workbook.Save JSON export C# | large CSV to JSON conversion Aspose.Cells | memory‑efficient CSV handling Aspose.Cells | C# Aspose.Cells data‑only JSON output

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Json;

// The example checks for the CSV file, loads it into an Aspose.Cells Workbook using LoadOptions with LoadFormat.Csv, and saves the workbook directly as a JSON file, providing a straightforward way to convert large CSV data to JSON in C# while keeping memory usage low.
class Program
{
    static void Main()
    {
        // Path to the large CSV file
        string csvPath = "input.csv";

        // Path where the JSON output will be saved
        string jsonPath = "output.json";

        // Verify that the input CSV file exists
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"Error: CSV file not found at '{csvPath}'.");
            return;
        }

        try
        {
            // Configure load options for CSV (use the non‑obsolete enum value)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

            // Load the CSV file into a workbook using the specified options.
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Save the workbook as JSON. Default options are sufficient for data‑only export.
            workbook.Save(jsonPath, SaveFormat.Json);

            Console.WriteLine($"JSON file successfully saved to '{jsonPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
