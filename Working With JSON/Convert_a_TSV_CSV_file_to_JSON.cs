using System;
using Aspose.Cells;               // Main Aspose.Cells namespace
using Aspose.Cells.Utility;      // For JsonSaveOptions (if needed)

class TsvToJsonConverter
{
    static void Main()
    {
        // Path to the source TSV (tab‑separated) file
        string tsvFilePath = "input.tsv";

        // Create a new workbook (in‑memory Excel file)
        Workbook workbook = new Workbook();

        // Access the first worksheet where the data will be imported
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Import the TSV data.
        //   splitter: "\t"  (tab character)
        //   convertNumericData: true  (numeric strings become numbers)
        //   firstRow / firstColumn: 0 (start at cell A1)
        cells.ImportCSV(tsvFilePath, "\t", true, 0, 0);

        // Configure JSON export options (optional – adjust as needed)
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export the data as a nested JSON structure
            ExportNestedStructure = true,
            // Include header row if the first row contains column names
            HasHeaderRow = true,
            // Indent the output for readability (4 spaces)
            Indent = "    "
        };

        // Save the workbook as a JSON file.
        // This uses Aspose.Cells' built‑in save mechanism, complying with the lifecycle rule.
        string jsonOutputPath = "output.json";
        workbook.Save(jsonOutputPath, jsonOptions);

        Console.WriteLine($"TSV data from '{tsvFilePath}' has been converted to JSON and saved as '{jsonOutputPath}'.");
    }
}