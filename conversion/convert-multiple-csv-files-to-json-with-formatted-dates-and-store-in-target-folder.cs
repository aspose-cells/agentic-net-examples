// Title: Convert multiple CSV files to JSON with custom date formatting using Aspose.Cells in C#
// AI Prompts: Write a C# console application that scans a folder for all *.csv files, loads each file with Aspose.Cells TxtLoadOptions (enabling date and numeric conversion), applies the custom date format "dd-MM-yyyy" to every DateTime cell, and saves the workbook as a .json file in a target directory using JsonSaveOptions. | Update the program to use a semicolon as the CSV delimiter and export dates in ISO 8601 format while preserving empty rows and cells in the generated JSON.
// Common Searches: how to convert many csv files to json with Aspose.Cells in C# | c# load csv with TxtLoadOptions and preserve date values | export Aspose.Cells workbook to json keeping empty rows | apply custom date format when saving json from csv using Aspose.Cells | set csv delimiter to semicolon and output iso 8601 dates in json
// Tags: batch csv to json conversion Aspose.Cells | date formatting dd-MM-yyyy Aspose.Cells | TxtLoadOptions enable date conversion C# | JsonSaveOptions include empty cells | apply style to DateTime cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace CsvToJsonConverter
{
    // A C# console program that iterates over every CSV file in a source directory, loads each file with Aspose.Cells TxtLoadOptions (converting dates and numbers), formats all DateTime cells to "dd-MM-yyyy", and saves the data as JSON files in a target folder using JsonSaveOptions that retain headers, empty cells, and native data types.
    class Program
    {
        static void Main()
        {
            // Folder containing source CSV files
            string sourceFolder = @"C:\Data\CsvFiles";
            // Folder where JSON files will be saved
            string targetFolder = @"C:\Data\JsonOutput";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure the target folder exists
            if (!Directory.Exists(targetFolder))
                Directory.CreateDirectory(targetFolder);

            // Get all CSV files in the source folder
            string[] csvFiles = Directory.GetFiles(sourceFolder, "*.csv");

            foreach (string csvPath in csvFiles)
            {
                // Guard against missing file (should not happen with GetFiles)
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"File not found: {csvPath}");
                    continue;
                }

                try
                {
                    // Load CSV into a workbook with date and numeric conversion enabled
                    TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
                    {
                        Separator = ',',               // CSV delimiter
                        ConvertDateTimeData = true,    // Convert date strings to DateTime
                        ConvertNumericData = true      // Convert numeric strings to numbers
                    };

                    Workbook workbook = new Workbook(csvPath, loadOptions);
                    Worksheet sheet = workbook.Worksheets[0];
                    Cells cells = sheet.Cells;

                    // Apply desired date format to all DateTime cells
                    string desiredDateFormat = "dd-MM-yyyy";
                    for (int row = 0; row <= cells.MaxDataRow; row++)
                    {
                        for (int col = 0; col <= cells.MaxDataColumn; col++)
                        {
                            Cell cell = cells[row, col];
                            if (cell.Type == CellValueType.IsDateTime)
                            {
                                Style style = cell.GetStyle();
                                // Use Custom format for date representation
                                style.Custom = desiredDateFormat;
                                cell.SetStyle(style);
                            }
                        }
                    }

                    // Prepare JSON save options
                    JsonSaveOptions jsonOptions = new JsonSaveOptions
                    {
                        HasHeaderRow = true,          // First row contains column names
                        ExportEmptyCells = true,      // Export empty cells as null
                        ExportAsString = false,       // Export values in native types
                        SkipEmptyRows = false         // Keep empty rows in output
                    };

                    // Build JSON file path (same name as CSV but .json extension)
                    string jsonFileName = Path.GetFileNameWithoutExtension(csvPath) + ".json";
                    string jsonPath = Path.Combine(targetFolder, jsonFileName);

                    // Save workbook as JSON using the configured options
                    workbook.Save(jsonPath, jsonOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{csvPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
