using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class CsvToJsonConverter
{
    static void Main()
    {
        try
        {
            // Folder containing source CSV files
            string sourceFolder = @"C:\CsvSource";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Folder where JSON files will be saved
            string targetFolder = @"C:\JsonTarget";

            // Ensure the target folder exists
            Directory.CreateDirectory(targetFolder);

            // Get all CSV files in the source folder
            string[] csvFiles = Directory.GetFiles(sourceFolder, "*.csv");

            foreach (string csvPath in csvFiles)
            {
                try
                {
                    // Verify CSV file exists (redundant but safe)
                    if (!File.Exists(csvPath))
                    {
                        Console.WriteLine($"File not found: {csvPath}");
                        continue;
                    }

                    // Create a new workbook (empty Excel file)
                    Workbook workbook = new Workbook();

                    // Configure load options for CSV import
                    TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
                    {
                        Separator = ',',
                        ConvertDateTimeData = true,
                        ConvertNumericData = true,
                        LoadStyleStrategy = TxtLoadStyleStrategy.ExactFormat
                    };

                    // Import CSV data into the first worksheet starting at cell A1
                    workbook.Worksheets[0].Cells.ImportCSV(csvPath, loadOptions, 0, 0);

                    // Apply a uniform date format to all cells that contain DateTime values
                    Worksheet sheet = workbook.Worksheets[0];
                    Cells cells = sheet.Cells;
                    Style dateStyle = workbook.CreateStyle();
                    dateStyle.Custom = "yyyy-MM-dd";

                    foreach (Cell cell in cells)
                    {
                        if (cell.Type == CellValueType.IsDateTime)
                        {
                            cell.SetStyle(dateStyle);
                        }
                    }

                    // Set JSON export options
                    JsonSaveOptions jsonOptions = new JsonSaveOptions
                    {
                        ExportEmptyCells = false,
                        HasHeaderRow = true,
                        ExportNestedStructure = false,
                        ExportAsString = false,
                        Indent = "  "
                    };

                    // Build the output JSON file path (same name as CSV, different extension)
                    string jsonFileName = Path.GetFileNameWithoutExtension(csvPath) + ".json";
                    string jsonPath = Path.Combine(targetFolder, jsonFileName);

                    // Save the workbook as a JSON file using the specified options
                    workbook.Save(jsonPath, jsonOptions);

                    Console.WriteLine($"Converted '{csvPath}' to '{jsonPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{csvPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}