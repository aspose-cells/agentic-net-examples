// Title: Batch convert CSV files in a folder to XLSX workbooks using Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells batch convert csv files in a directory to xlsx | how to load multiple csv files and save as Excel workbooks with Aspose.Cells | process all csv files in a folder and export to .xlsx using Aspose.Cells C# | Aspose.Cells loop over csv files and generate xlsx files with error handling
// Tags: batch csv to xlsx conversion Aspose.Cells | load csv workbook with LoadOptions C# | save workbook as xlsx format Aspose.Cells | iterate directory files for Excel export C# | error handling during bulk csv conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace CsvToXlsxBatchProcessor
{
    // The example scans a specified input folder for *.csv files, loads each file into an Aspose.Cells Workbook using default CSV load options, and saves it as an .xlsx file in an output folder, creating the folder if needed and handling missing files or processing errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Input directory containing CSV files
            string inputDirectory = args.Length > 0 ? args[0] : @"C:\InputCsv";
            // Output directory for generated XLSX files
            string outputDirectory = args.Length > 1 ? args[1] : @"C:\OutputXlsx";

            // Ensure input directory exists
            if (!Directory.Exists(inputDirectory))
            {
                Console.WriteLine($"Input directory does not exist: {inputDirectory}");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            // Get all CSV files in the input directory
            string[] csvFiles = Directory.GetFiles(inputDirectory, "*.csv", SearchOption.TopDirectoryOnly);

            foreach (string csvFilePath in csvFiles)
            {
                // Verify the file exists before processing
                if (!File.Exists(csvFilePath))
                {
                    Console.WriteLine($"File not found: {csvFilePath}");
                    continue;
                }

                try
                {
                    // Load CSV using default options (comma delimiter)
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

                    // Load the CSV file into a workbook
                    Workbook workbook = new Workbook(csvFilePath, loadOptions);

                    // Determine the output XLSX file path
                    string outputFileName = Path.GetFileNameWithoutExtension(csvFilePath) + ".xlsx";
                    string outputFilePath = Path.Combine(outputDirectory, outputFileName);

                    // Save the workbook as XLSX
                    workbook.Save(outputFilePath, SaveFormat.Xlsx);

                    Console.WriteLine($"Converted '{Path.GetFileName(csvFilePath)}' to '{outputFileName}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{csvFilePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
