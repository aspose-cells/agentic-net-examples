using System;
using System.IO;
using Aspose.Cells;

class CsvToJsonStreaming
{
    static void Main()
    {
        // Paths for input CSV and output JSON
        string csvPath = "large_input.csv";
        string jsonPath = "output.json";

        // Ensure the CSV file exists; create a sample if it does not
        if (!File.Exists(csvPath))
        {
            File.WriteAllText(csvPath,
                "Id,Name,Age\n" +
                "1,John Doe,30\n" +
                "2,Jane Smith,25\n" +
                "3,Bob Johnson,40");
        }

        // Load the CSV file into a workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.CSV);
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportNestedStructure = true,
            SkipEmptyRows = true,
            HasHeaderRow = true,
            ExportEmptyCells = false,
            ExportAsString = false
        };

        // Save the workbook as JSON
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine("CSV has been successfully converted to JSON.");
    }
}