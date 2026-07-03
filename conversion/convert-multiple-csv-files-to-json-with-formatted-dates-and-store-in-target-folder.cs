using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class CsvToJsonConverter
{
    static void Main()
    {
        // Folder containing source CSV files
        string sourceFolder = @"C:\Data\CsvFiles";
        // Folder where JSON files will be saved
        string targetFolder = @"C:\Data\JsonOutput";

        // Ensure target folder exists
        Directory.CreateDirectory(targetFolder);

        // Process each CSV file in the source folder
        foreach (string csvPath in Directory.GetFiles(sourceFolder, "*.csv"))
        {
            // Load CSV with options that convert date strings to DateTime values
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
            loadOptions.ConvertDateTimeData = true;          // enable date conversion
            loadOptions.ConvertNumericData = true;           // enable numeric conversion
            loadOptions.LoadStyleStrategy = TxtLoadStyleStrategy.ExactFormat; // keep original format

            // Load the CSV into a workbook
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Apply a uniform date format to all date cells (optional)
            // This sets the custom number format for the default style; date cells will inherit it.
            workbook.DefaultStyle.Custom = "yyyy-MM-dd";

            // Prepare JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export as a JSON object (not an array) even if only one worksheet exists
                AlwaysExportAsJsonObject = true,
                // Include header row if present
                HasHeaderRow = true,
                // Export empty cells as null (can be changed as needed)
                ExportEmptyCells = true,
                // Do not export as nested structure for flat CSV data
                ExportNestedStructure = false,
                // Export cell values as strings to preserve formatting
                ExportAsString = true,
                // Indent JSON for readability
                Indent = "  "
            };

            // Determine output JSON file path
            string jsonFileName = Path.GetFileNameWithoutExtension(csvPath) + ".json";
            string jsonPath = Path.Combine(targetFolder, jsonFileName);

            // Save workbook as JSON
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Converted '{Path.GetFileName(csvPath)}' to JSON at '{jsonPath}'.");
        }
    }
}