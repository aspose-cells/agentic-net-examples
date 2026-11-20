// Title: Batch convert CSV files to JSON with ISO‑8601 dates using Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells convert folder of CSV files to JSON | how to format date cells as yyyy-MM-dd when saving CSV to JSON with Aspose.Cells | batch process multiple CSV files to JSON using Aspose.Cells LoadOptions | Aspose.Cells C# convert CSV to JSON and change date format | save CSV workbook as JSON with custom date format using Aspose.Cells
// Tags: Aspose.Cells CSV to JSON batch conversion | C# format Excel date cells ISO 8601 | LoadOptions LoadFormat.Csv Aspose.Cells | JsonSaveOptions save workbook as JSON | process multiple files directory Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program scans a source directory for .csv files, loads each file into an Aspose.Cells Workbook using LoadFormat.Csv, converts any DateTime cells to ISO‑8601 (yyyy-MM-dd) strings, and saves the result as a .json file in a target folder with JsonSaveOptions.
class CsvToJsonConverter
{
    // Source folder containing CSV files
    private const string SourceFolder = @"C:\InputCsv";
    // Destination folder for JSON files
    private const string TargetFolder = @"C:\OutputJson";

    static void Main()
    {
        try
        {
            // Ensure source folder exists
            if (!Directory.Exists(SourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {SourceFolder}");
                return;
            }

            // Ensure target folder exists
            Directory.CreateDirectory(TargetFolder);

            // Get all CSV files in the source folder
            string[] csvFiles = Directory.GetFiles(SourceFolder, "*.csv");

            foreach (string csvPath in csvFiles)
            {
                try
                {
                    // Verify the CSV file still exists before loading
                    if (!File.Exists(csvPath))
                    {
                        Console.WriteLine($"File not found: {csvPath}");
                        continue;
                    }

                    // Load CSV into a Workbook
                    Workbook workbook = new Workbook(csvPath, new LoadOptions(LoadFormat.Csv));

                    // Process the first worksheet (CSV files usually have a single sheet)
                    Worksheet sheet = workbook.Worksheets[0];
                    Cells cells = sheet.Cells;

                    // Iterate through all used cells
                    foreach (Cell cell in cells)
                    {
                        // Check if the cell contains a DateTime value
                        if (cell.Type == CellValueType.IsDateTime)
                        {
                            // Format the date as ISO string (yyyy-MM-dd) and store as string
                            DateTime dt = cell.DateTimeValue;
                            cell.PutValue(dt.ToString("yyyy-MM-dd"));
                        }
                    }

                    // Prepare JSON save options (default options are sufficient)
                    JsonSaveOptions jsonOptions = new JsonSaveOptions();

                    // Build output JSON file path
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(csvPath);
                    string jsonPath = Path.Combine(TargetFolder, fileNameWithoutExt + ".json");

                    // Save the workbook as JSON
                    workbook.Save(jsonPath, jsonOptions);

                    Console.WriteLine($"Converted: {csvPath} -> {jsonPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{csvPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Conversion completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
