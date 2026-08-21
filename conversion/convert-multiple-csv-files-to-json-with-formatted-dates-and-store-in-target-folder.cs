// Title: Batch convert CSV to JSON with ISO‑8601 dates using Aspose.Cells for .NET
// Description: Scans a source folder for *.csv files, loads each with TxtLoadOptions (auto‑convert dates and numbers), applies a uniform "yyyy-MM-dd" style to every DateTime cell, and saves the workbook as a flat JSON file (header row kept, empty cells omitted) in a target directory.
// Keywords: Aspose.Cells | C# | CSV to JSON conversion | batch processing | date formatting ISO 8601 | TxtLoadOptions ConvertDateTimeData | JsonSaveOptions | folder based conversion | skip empty cells | flat JSON structure
// Common Searches: C# batch convert csv files to json aspocells | Aspose.Cells export csv with custom date format | how to apply yyyy-MM-dd to dates when saving json | convert multiple csv to json in a folder .NET | skip empty cells Aspose.Cells JsonSaveOptions
// Developer Intent: Read each CSV in a directory, enforce a consistent date format, and write the result as JSON to a specified output folder.
// Use Cases: Convert daily CSV logs into ISO‑8601 JSON for API ingestion while preserving numeric types. | Generate configuration JSON from CSV templates, automatically formatting dates and removing blanks. | Migrate legacy CSV datasets to NoSQL stores by producing flat JSON files with uniform date representation.
// AI Prompts: Write a reusable C# method that takes source and target folder paths and uses Aspose.Cells to batch‑convert CSV files to JSON with a custom date format. | Explain how TxtLoadOptions.ConvertDateTimeData and JsonSaveOptions.HasHeaderRow affect the structure and content of the exported JSON. | Suggest robust error‑handling and logging strategies for the conversion loop, including how to capture files that fail and why.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace CsvToJsonConverter
{
    // Scans a source folder for *.csv files, loads each with TxtLoadOptions (auto‑convert dates and numbers), applies a uniform "yyyy-MM-dd" style to every DateTime cell, and saves the workbook as a flat JSON file (header row kept, empty cells omitted) in a target directory.
    class Program
    {
        static void Main()
        {
            // Folder containing source CSV files
            string sourceFolder = @"C:\SourceCsv";

            // Folder where JSON files will be saved
            string targetFolder = @"C:\TargetJson";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the target folder exists
            Directory.CreateDirectory(targetFolder);

            try
            {
                // Process each CSV file in the source folder
                foreach (string csvPath in Directory.GetFiles(sourceFolder, "*.csv"))
                {
                    try
                    {
                        // Determine output JSON file name
                        string jsonFileName = Path.GetFileNameWithoutExtension(csvPath) + ".json";
                        string jsonPath = Path.Combine(targetFolder, jsonFileName);

                        // Load options for CSV: convert dates and numbers automatically
                        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
                        {
                            ConvertDateTimeData = true,
                            ConvertNumericData = true
                        };

                        // Load the CSV file into a workbook
                        Workbook workbook = new Workbook(csvPath, loadOptions);

                        // Apply a uniform date format to all cells that contain DateTime values
                        Worksheet sheet = workbook.Worksheets[0];
                        Cells cells = sheet.Cells;
                        for (int row = 0; row <= cells.MaxDataRow; row++)
                        {
                            for (int col = 0; col <= cells.MaxDataColumn; col++)
                            {
                                Cell cell = cells[row, col];
                                if (cell.Type == CellValueType.IsDateTime)
                                {
                                    Style style = cell.GetStyle();
                                    style.Custom = "yyyy-MM-dd"; // Desired date format
                                    cell.SetStyle(style);
                                }
                            }
                        }

                        // Configure JSON save options
                        JsonSaveOptions jsonOptions = new JsonSaveOptions
                        {
                            HasHeaderRow = true,          // First row contains column names
                            ExportEmptyCells = false,     // Skip empty cells
                            ExportNestedStructure = false // Flat structure
                        };

                        // Save the workbook as a JSON file
                        workbook.Save(jsonPath, jsonOptions);

                        Console.WriteLine($"Converted '{Path.GetFileName(csvPath)}' to '{jsonFileName}'.");
                    }
                    catch (Exception exFile)
                    {
                        Console.WriteLine($"Error processing file '{csvPath}': {exFile.Message}");
                    }
                }

                Console.WriteLine("All files have been processed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
