// Title: Convert multiple CSV files to ODS in C# with Aspose.Cells and verify each output
// Common Searches: C# code to convert all CSV files in a folder to ODS using Aspose.Cells | how to validate ODS files after conversion from CSV with Aspose.Cells .NET | batch processing CSV to ODS with default ODF version in Aspose.Cells | Aspose.Cells example for loading CSV and saving as ODS in .NET | verify ODS compliance by reloading saved file in C# Aspose.Cells
// Tags: batch CSV to ODS conversion Aspose.Cells | load CSV workbook LoadOptions C# | save workbook as ODS OdsSaveOptions | verify ODS file by reloading Aspose.Cells | default ODF version handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace CsvToOdsConverter
{
    // The program enumerates CSV files in a specified directory, loads each file into an Aspose.Cells Workbook using LoadOptions for CSV, saves it as an ODS file with default OdsSaveOptions, and then reloads the saved ODS to confirm successful conversion and ODF compliance.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing CSV files
            string inputDirectory = @"C:\InputCsv";
            // Directory where ODS files will be saved
            string outputDirectory = @"C:\OutputOds";

            // Verify input directory exists
            if (!Directory.Exists(inputDirectory))
            {
                Console.WriteLine($"Input directory does not exist: {inputDirectory}");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            // Get all CSV files in the input directory
            string[] csvFiles = Directory.GetFiles(inputDirectory, "*.csv", SearchOption.TopDirectoryOnly);

            foreach (string csvPath in csvFiles)
            {
                try
                {
                    // Verify the CSV file exists before loading
                    if (!File.Exists(csvPath))
                    {
                        Console.WriteLine($"CSV file not found: {csvPath}");
                        continue;
                    }

                    // Load CSV file using default CSV load options
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
                    Workbook workbook = new Workbook(csvPath, loadOptions);

                    // Prepare ODS save options (default ODF version is used automatically)
                    OdsSaveOptions odsOptions = new OdsSaveOptions();

                    // Determine output file name with .ods extension
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(csvPath);
                    string odsPath = Path.Combine(outputDirectory, fileNameWithoutExt + ".ods");

                    // Save workbook as ODS
                    workbook.Save(odsPath, odsOptions);

                    // Verify compliance by attempting to reload the saved ODS file
                    if (File.Exists(odsPath))
                    {
                        // If the file loads without exception, we consider it compliant
                        Workbook verifyWorkbook = new Workbook(odsPath);
                        Console.WriteLine($"Successfully converted and verified: {Path.GetFileName(csvPath)} -> {Path.GetFileName(odsPath)}");
                    }
                    else
                    {
                        Console.WriteLine($"Failed to create ODS file for: {Path.GetFileName(csvPath)}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing {Path.GetFileName(csvPath)}: {ex.Message}");
                }
            }

            Console.WriteLine("Conversion process completed.");
        }
    }
}
