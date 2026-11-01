// Title: Convert a folder of CSV files to ODS spreadsheets with Aspose.Cells for .NET and verify each output
// Common Searches: c# Aspose.Cells convert all csv files in a directory to ods | batch convert csv to ods using Aspose.Cells .NET example | how to validate ods files after saving with Aspose.Cells in C# | load csv with LoadOptions and save as ODS default ODF version Aspose.Cells | process multiple csv files and generate ods output using Aspose.Cells
// Tags: Aspose.Cells batch CSV to ODS conversion | C# LoadOptions CSV Aspose.Cells | SaveFormat.ODS default ODF version | verify ODS file existence C# | process multiple CSV files Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program scans a specified input folder for *.csv files, loads each file into an Aspose.Cells Workbook using CSV LoadOptions, saves it as an ODS file with the default ODF version in an output folder, and then reloads the saved ODS file to confirm successful conversion, reporting the result for each file.
class CsvToOdsConverter
{
    static void Main(string[] args)
    {
        // Directory containing CSV files
        string inputFolder = @"C:\InputCsv";
        // Directory where ODS files will be saved
        string outputFolder = @"C:\OutputOds";

        // Verify input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input directory does not exist: {inputFolder}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        string[] csvFiles;
        try
        {
            // Get all CSV files in the input folder
            csvFiles = Directory.GetFiles(inputFolder, "*.csv", SearchOption.TopDirectoryOnly);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error accessing input directory: {ex.Message}");
            return;
        }

        foreach (string csvPath in csvFiles)
        {
            try
            {
                // Load CSV file into a Workbook
                var loadOptions = new LoadOptions(LoadFormat.Csv);
                var workbook = new Workbook(csvPath, loadOptions);

                // Prepare ODS output path (same name, different extension)
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(csvPath);
                string odsPath = Path.Combine(outputFolder, fileNameWithoutExt + ".ods");

                // Save workbook as ODS
                workbook.Save(odsPath, SaveFormat.ODS);

                // Verify compliance: check that the file exists and can be opened as ODS
                if (File.Exists(odsPath))
                {
                    var verifyOptions = new LoadOptions(LoadFormat.ODS);
                    var verifyWorkbook = new Workbook(odsPath, verifyOptions);
                    Console.WriteLine($"Successfully converted '{Path.GetFileName(csvPath)}' to ODS.");
                }
                else
                {
                    Console.WriteLine($"Failed to create ODS file for '{Path.GetFileName(csvPath)}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{Path.GetFileName(csvPath)}': {ex.Message}");
            }
        }

        Console.WriteLine("Conversion process completed.");
    }
}
