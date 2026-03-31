using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class CsvToJsonConverter
{
    static void Main()
    {
        // Input CSV file path (replace with your actual file location)
        string csvFilePath = "input.csv";

        // Output JSON file path
        string jsonFilePath = "output.json";

        // Create a new workbook (empty Excel file)
        Workbook workbook = new Workbook();

        // Import the CSV data into the first worksheet.
        // Parameters: file name, delimiter, convert numeric strings, start row, start column.
        workbook.Worksheets[0].Cells.ImportCSV(csvFilePath, ",", true, 0, 0);

        // Configure JSON save options to preserve the Excel structural hierarchy
        // and retain original cell types (numbers, dates, strings, etc.).
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ToExcelStruct = true,          // Export as Excel structural JSON
            ExportEmptyCells = true,       // Include empty cells as null in JSON
            ExportAsString = false,        // Keep original data types
            HasHeaderRow = false           // No special header handling (adjust if needed)
        };

        // Save the workbook directly as a JSON file using the configured options.
        workbook.Save(jsonFilePath, jsonOptions);
    }
}