using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class PdfEmbeddedCsvToJson
{
    static void Main()
    {
        // Path to the CSV file that contains the embedded data.
        string csvFilePath = Path.Combine(Environment.CurrentDirectory, "embedded_dataset.csv");

        // Ensure the CSV file exists; create a sample if it does not.
        if (!File.Exists(csvFilePath))
        {
            File.WriteAllText(csvFilePath,
                "Id,Name,Value\n" +
                "1,Alpha,100\n" +
                "2,Beta,200\n" +
                "3,Gamma,300");
        }

        // Path where the resulting JSON will be saved.
        string jsonOutputPath = Path.Combine(Environment.CurrentDirectory, "dataset.json");

        // Create a new empty workbook.
        Workbook workbook = new Workbook();

        // Access the first worksheet.
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Import the CSV data into the worksheet.
        cells.ImportCSV(csvFilePath, ",", true, 0, 0);

        // Determine the used range dimensions.
        int totalRows = cells.MaxDataRow + 1;      // MaxDataRow is zero‑based.
        int totalColumns = cells.MaxDataColumn + 1;

        // Create a range that covers all imported data.
        Aspose.Cells.Range dataRange = cells.CreateRange(0, 0, totalRows, totalColumns);

        // Configure JSON export options.
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,
            ExportEmptyCells = true,
            ExportNestedStructure = false
        };

        // Export the range to a JSON string.
        string jsonResult = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

        // Write the JSON string to the output file.
        File.WriteAllText(jsonOutputPath, jsonResult);

        Console.WriteLine("CSV data successfully converted to JSON.");
        Console.WriteLine($"JSON saved to: {jsonOutputPath}");
    }
}