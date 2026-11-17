// Title: Load a CSV file into an Aspose.Cells Workbook and export it as JSON using C#
// Common Searches: C# code sample for loading a CSV file with Aspose.Cells LoadOptions | How to export an Aspose.Cells workbook to JSON in .NET | Convert CSV data to JSON format using Aspose.Cells in C# | Aspose.Cells CSV import and JSON export example for .NET developers | Saving a CSV‑loaded workbook as JSON with Aspose.Cells C# API
// Tags: Aspose.Cells CSV LoadOptions C# | Aspose.Cells SaveFormat JSON C# | Aspose.Cells workbook CSV import | Aspose.Cells workbook JSON export | C# conversion of CSV workbook to JSON using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the CSV file exists, loads it into an Aspose.Cells Workbook using LoadOptions with LoadFormat.Csv, then saves the workbook as a JSON file via SaveFormat.Json, handling any exceptions that may occur.
class CsvToJsonConverter
{
    static void Main()
    {
        // Path to the source CSV file
        string csvFilePath = "input.csv";

        // Verify that the CSV file exists to avoid FileNotFoundException
        if (!File.Exists(csvFilePath))
        {
            Console.WriteLine($"Error: CSV file '{csvFilePath}' not found.");
            return;
        }

        try
        {
            // Load the CSV file into an Aspose.Cells Workbook using CSV load options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvFilePath, loadOptions);

            // Define the output JSON file path
            string jsonFilePath = "output.json";

            // Save the workbook as JSON
            workbook.Save(jsonFilePath, SaveFormat.Json);

            // Inform the user that the conversion is complete
            Console.WriteLine($"CSV file '{csvFilePath}' has been converted to JSON and saved as '{jsonFilePath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
