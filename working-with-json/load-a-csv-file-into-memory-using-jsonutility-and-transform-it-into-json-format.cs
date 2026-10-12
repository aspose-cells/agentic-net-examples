// Title: Read a CSV file with Aspose.Cells, convert it to a DataTable, and save as indented JSON using System.Text.Json (C#)
// AI Prompts: Write C# code that loads a CSV file into an Aspose.Cells Workbook, exports the worksheet to a DataTable, and writes the data as pretty‑printed JSON with System.Text.Json. | Enhance the program to accept input and output file paths as command‑line arguments and add optional parameters for JSON indentation and culture‑specific number formatting.
// Common Searches: c# aspocells read csv and export to json file | how to convert csv to json using Aspose.Cells workbook | serialize datatable to formatted json with System.Text.Json c# | aspocells csv to json example .net core | command line csv to json conversion aspocells c#
// Tags: Aspose.Cells CSV import to DataTable | System.Text.Json DataTable serialization | CSV to JSON conversion C# | Workbook CSV loading Aspose.Cells | Indented JSON output .NET

using System;
using System.IO;
using System.Data;
using System.Text.Json;
using Aspose.Cells;

// The example verifies the CSV file exists, loads it into an Aspose.Cells Workbook, accesses the first worksheet, determines the used range, exports the data to a DataTable, and then serializes that DataTable to indented JSON using System.Text.Json. The JSON is printed to the console and written to an output file, with error handling for unexpected exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.csv";
            const string outputPath = "output.json";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the CSV file into a workbook (Aspose.Cells can read CSV directly)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet which contains the CSV data
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range dimensions
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalColumns = sheet.Cells.MaxDataColumn + 1;

            // Export the worksheet data to a DataTable
            DataTable dt = sheet.Cells.ExportDataTable(0, 0, totalRows, totalColumns, true);

            // Serialize the DataTable to a JSON string with indentation
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(dt, jsonOptions);

            // Output the JSON to console and file
            Console.WriteLine(json);
            File.WriteAllText(outputPath, json);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
