// Title: Convert multiple CSV files to JSON with custom yyyy-MM-dd date formatting using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells batch convert CSV files to JSON with specific date format | How to apply custom date format to cells before saving workbook as JSON in Aspose.Cells | Convert all CSV files in a directory to JSON using Aspose.Cells and preserve date format | Aspose.Cells load CSV, format DateTime cells to yyyy-MM-dd, export to JSON | C# script to process multiple CSVs and output JSON with formatted dates using Aspose.Cells
// Tags: batch CSV to JSON conversion Aspose.Cells C# | custom date format cells Aspose.Cells | load CSV workbook LoadOptions Aspose.Cells | save workbook as JSON Aspose.Cells | iterate worksheets cells format dates C#

using System;
using System.IO;
using Aspose.Cells;

namespace CsvToJsonConverter
{
    // The C# program scans a source folder for CSV files, loads each file into an Aspose.Cells Workbook using LoadOptions, iterates through all worksheets and cells to apply a 'yyyy-MM-dd' custom format to any DateTime values, and then saves each workbook as a JSON file in a target directory, creating the target folder if needed.
    class Program
    {
        static void Main(string[] args)
        {
            // Define source folder containing CSV files and target folder for JSON output
            string sourceFolder = @"C:\SourceCsv";
            string targetFolder = @"C:\TargetJson";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure the target directory exists
            if (!Directory.Exists(targetFolder))
                Directory.CreateDirectory(targetFolder);

            string[] csvFiles;
            try
            {
                // Get all CSV files in the source folder
                csvFiles = Directory.GetFiles(sourceFolder, "*.csv");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to enumerate CSV files: {ex.Message}");
                return;
            }

            foreach (string csvFilePath in csvFiles)
            {
                try
                {
                    // Load CSV file into a workbook
                    var loadOptions = new LoadOptions(LoadFormat.Csv);
                    var workbook = new Workbook(csvFilePath, loadOptions);

                    // Iterate through all worksheets and cells to format date values
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        Cells cells = sheet.Cells;
                        foreach (Cell cell in cells)
                        {
                            // Check if the cell contains a DateTime value
                            if (cell.Type == CellValueType.IsDateTime)
                            {
                                // Apply a custom date format (e.g., yyyy-MM-dd)
                                Style style = cell.GetStyle();
                                style.Custom = "yyyy-MM-dd";
                                cell.SetStyle(style);
                            }
                        }
                    }

                    // Prepare JSON output path
                    string jsonFileName = Path.GetFileNameWithoutExtension(csvFilePath) + ".json";
                    string jsonFilePath = Path.Combine(targetFolder, jsonFileName);

                    // Save workbook as JSON
                    var jsonOptions = new JsonSaveOptions();
                    workbook.Save(jsonFilePath, jsonOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{csvFilePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Conversion completed.");
        }
    }
}
