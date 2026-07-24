// Title: Batch Convert CSV to JSON with Exact Date Formatting using Aspose.Cells for .NET (C#)
// Description: C# utility that scans a source folder, loads each *.csv file into an Aspose.Cells workbook with TxtLoadOptions (ConvertDateTimeData, ConvertNumericData, ExactFormat), and exports the used range to a JSON file. The JSON preserves the original date strings, includes a header row, skips empty cells, and is saved to a target directory with the same base name.
// Keywords: Aspose.Cells CSV to JSON | C# batch CSV conversion | ExactFormat date preservation | TxtLoadOptions ConvertDateTimeData | .NET CSV to JSON utility | JsonSaveOptions header row | folder based CSV processing | Aspose.Cells JsonExport
// Common Searches: How to convert multiple CSV files to JSON with Aspose.Cells | C# preserve original date format when exporting CSV to JSON | Batch CSV to JSON conversion using Aspose.Cells .NET | Export CSV data to JSON with header row and no empty cells | Aspose.Cells ExactFormat example C#
// Developer Intent: The developer needs a reliable way to batch‑process CSV files into JSON while keeping the original date representation and proper data types.
// Use Cases: Convert daily CSV logs into JSON payloads for a REST API without altering date strings. | Migrate legacy configuration CSVs to JSON for a .NET microservice, ensuring numeric and date values are correctly typed. | Generate front‑end friendly JSON reports from CSV datasets, exporting only populated cells with column headers.
// AI Prompts: Write C# code that uses Aspose.Cells to read a CSV, keep its date format, and output JSON with a header row. | Show how to add robust error handling and logging for a batch CSV‑to‑JSON conversion using Aspose.Cells. | Demonstrate modifying the export to produce a nested JSON structure instead of a flat array with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

// C# utility that scans a source folder, loads each *.csv file into an Aspose.Cells workbook with TxtLoadOptions (ConvertDateTimeData, ConvertNumericData, ExactFormat), and exports the used range to a JSON file. The JSON preserves the original date strings, includes a header row, skips empty cells, and is saved to a target directory with the same base name.
public class CsvToJsonConverter
{
    // Converts all CSV files in a source folder to JSON files in a target folder.
    // Dates are preserved using ExactFormat style so the original format appears in JSON.
    public static void ConvertFolder(string sourceFolder, string targetFolder)
    {
        try
        {
            // Verify source folder exists.
            if (!Directory.Exists(sourceFolder))
                throw new DirectoryNotFoundException($"Source folder not found: {sourceFolder}");

            // Ensure the target directory exists.
            if (!Directory.Exists(targetFolder))
                Directory.CreateDirectory(targetFolder);

            // Get all CSV files in the source folder.
            string[] csvFiles = Directory.GetFiles(sourceFolder, "*.csv", SearchOption.TopDirectoryOnly);

            foreach (string csvPath in csvFiles)
            {
                try
                {
                    // Create a new workbook and get its first worksheet.
                    Workbook workbook = new Workbook();
                    Worksheet worksheet = workbook.Worksheets[0];
                    Cells cells = worksheet.Cells;

                    // Configure load options for CSV import.
                    TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
                    {
                        ConvertDateTimeData = true,   // Convert date strings to DateTime.
                        ConvertNumericData = true,    // Convert numeric strings to numbers.
                        LoadStyleStrategy = TxtLoadStyleStrategy.ExactFormat // Preserve original format.
                    };

                    // Import the CSV data starting at cell A1 (row 0, column 0).
                    cells.ImportCSV(csvPath, loadOptions, 0, 0);

                    // Determine the used range of the worksheet.
                    int maxRow = cells.MaxDataRow;
                    int maxColumn = cells.MaxDataColumn;
                    AsposeRange usedRange = cells.CreateRange(0, 0, maxRow + 1, maxColumn + 1);

                    // Configure JSON export options.
                    JsonSaveOptions jsonOptions = new JsonSaveOptions
                    {
                        HasHeaderRow = true,          // First row contains column names.
                        ExportEmptyCells = false,     // Do not include empty cells.
                        ExportNestedStructure = false // Flat structure.
                    };

                    // Export the range to a JSON string.
                    string json = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

                    // Write the JSON string to a file with the same base name as the CSV.
                    string jsonFileName = Path.GetFileNameWithoutExtension(csvPath) + ".json";
                    string jsonPath = Path.Combine(targetFolder, jsonFileName);
                    File.WriteAllText(jsonPath, json);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{csvPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Conversion failed: {ex.Message}");
        }
    }

    // Example usage.
    public static void Main()
    {
        string sourceFolder = @"C:\Data\CsvFiles";
        string targetFolder = @"C:\Data\JsonOutput";

        ConvertFolder(sourceFolder, targetFolder);

        Console.WriteLine("Conversion completed.");
    }
}
