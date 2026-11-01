// Title: Convert multiple CSV files with a custom delimiter to XLSX in C# using Aspose.Cells
// Common Searches: C# Aspose.Cells load CSV files with semicolon delimiter and save as XLSX | How to batch convert all CSV files in a folder to Excel using Aspose.Cells | Aspose.Cells TxtLoadOptions custom separator for CSV to XLSX conversion in .NET | Process directory of CSV files and generate XLSX workbooks with Aspose.Cells | Convert CSV files with custom delimiter to Excel programmatically in C#
// Tags: Aspose.Cells TxtLoadOptions custom delimiter | batch CSV to XLSX conversion C# | load CSV with semicolon separator Aspose.Cells | save workbook as XLSX programmatically | iterate over directory files Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a specified input folder for all *.csv files, uses Aspose.Cells TxtLoadOptions to define a custom delimiter (e.g., semicolon), loads each CSV into a Workbook, and saves it as an XLSX file in an output folder. It creates the output directory if needed, logs each conversion, and handles errors gracefully.
class Program
{
    static void Main(string[] args)
    {
        // Folder containing the source CSV files
        string inputFolder = @"C:\InputCsv";

        // Folder where the generated XLSX files will be saved
        string outputFolder = @"C:\OutputXlsx";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Define the custom delimiter used in the CSV files (e.g., semicolon)
        char customDelimiter = ';';

        // Retrieve all CSV files from the input directory
        string[] csvFiles = Directory.GetFiles(inputFolder, "*.csv");

        foreach (string csvPath in csvFiles)
        {
            try
            {
                // Verify the CSV file still exists
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"File not found: {csvPath}");
                    continue;
                }

                // Configure load options for CSV with the custom delimiter
                TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
                {
                    Separator = customDelimiter
                };

                // Load the CSV file into a workbook using the specified options
                Workbook workbook = new Workbook(csvPath, loadOptions);

                // Construct the output XLSX file path (same name, different extension)
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(csvPath);
                string xlsxPath = Path.Combine(outputFolder, fileNameWithoutExt + ".xlsx");

                // Save the workbook as an XLSX file
                workbook.Save(xlsxPath, SaveFormat.Xlsx);

                Console.WriteLine($"Converted: {csvPath} -> {xlsxPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{csvPath}': {ex.Message}");
            }
        }
    }
}
