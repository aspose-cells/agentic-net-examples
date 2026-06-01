using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace CsvBatchToXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing source CSV files
            string sourceDirectory = @"C:\Data\CsvFiles";

            // Directory where the generated XLSX files will be saved
            string outputDirectory = @"C:\Data\XlsxOutput";

            // Ensure the output directory exists
            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            // Custom delimiter to be used while reading CSV files (e.g., semicolon)
            char customDelimiter = ';';

            // Get all CSV files in the source directory
            string[] csvFiles = Directory.GetFiles(sourceDirectory, "*.csv");

            foreach (string csvPath in csvFiles)
            {
                try
                {
                    // Prepare load options for CSV with the custom delimiter
                    TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
                    loadOptions.Separator = customDelimiter;          // Set custom delimiter
                    loadOptions.ConvertNumericData = true;            // Optional: convert numbers automatically

                    // Prepare save options for XLSX format
                    OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();

                    // Determine the output XLSX file path
                    string xlsxFileName = Path.GetFileNameWithoutExtension(csvPath) + ".xlsx";
                    string xlsxPath = Path.Combine(outputDirectory, xlsxFileName);

                    // Convert CSV to XLSX using the provided ConversionUtility method
                    ConversionUtility.Convert(csvPath, loadOptions, xlsxPath, saveOptions);

                    Console.WriteLine($"Converted: {Path.GetFileName(csvPath)} -> {xlsxFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting file '{csvPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}