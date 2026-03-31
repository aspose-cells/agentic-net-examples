using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

class CsvToJsonConverter
{
    static void Main()
    {
        // Paths for source CSV, intermediate XLSX, and final JSON output
        string csvPath = "input.csv";
        string xlsxPath = "intermediate.xlsx";
        string jsonPath = "output.json";

        // Ensure the CSV file exists (create a sample if needed)
        if (!File.Exists(csvPath))
        {
            File.WriteAllText(csvPath,
                "Name,Age,City\nJohn,30,New York\nAlice,25,London\nBob,35,Paris");
        }

        // Convert CSV to XLSX using the provided ConversionUtility rule
        ConversionUtility.Convert(csvPath, xlsxPath);

        // Load the converted workbook
        Workbook workbook = new Workbook(xlsxPath);
        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the used range (all rows and columns with data)
        int lastRow = worksheet.Cells.MaxDataRow;          // zero‑based index of last data row
        int lastColumn = worksheet.Cells.MaxDataColumn;    // zero‑based index of last data column

        // Create a range that covers the entire used area
        AsposeRange usedRange = worksheet.Cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

        // Configure JSON export options to preserve all cells
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportEmptyCells = true,   // keep empty cells as null
            HasHeaderRow = false,      // treat first row as data (preserve as is)
            SkipEmptyRows = false      // do not skip empty rows
        };

        // Export the range to a JSON string using the provided ExportRangeToJson rule
        string jsonContent = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

        // Write the JSON string to the output file
        File.WriteAllText(jsonPath, jsonContent);

        // Optional cleanup of the intermediate XLSX file
        if (File.Exists(xlsxPath))
        {
            File.Delete(xlsxPath);
        }

        Console.WriteLine($"CSV data has been converted to JSON and saved to '{jsonPath}'.");
    }
}