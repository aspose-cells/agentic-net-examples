using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class DifJsonToCsvConverter
{
    static void Main(string[] args)
    {
        // Input JSON file containing DIF data (replace with actual path or pass as argument)
        string jsonFilePath = args.Length > 0 ? args[0] : "input.json";

        // Read the entire JSON content
        string jsonContent = File.ReadAllText(jsonFilePath);

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Configure JSON import options:
        // - Treat JSON arrays as tables so that columns align correctly
        // - Convert numeric and date strings to proper cell types
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true,
            ConvertNumericOrDate = true
        };

        // Import JSON data starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, cells, 0, 0, layoutOptions);

        // Optional: auto‑fit columns for better CSV appearance
        worksheet.AutoFitColumns();

        // Determine output CSV file name (same name, .csv extension)
        string csvFilePath = Path.ChangeExtension(jsonFilePath, ".csv");

        // Save the workbook as CSV; Aspose.Cells automatically handles column alignment
        workbook.Save(csvFilePath, SaveFormat.Csv);

        Console.WriteLine($"Conversion completed. CSV saved to: {csvFilePath}");
    }
}