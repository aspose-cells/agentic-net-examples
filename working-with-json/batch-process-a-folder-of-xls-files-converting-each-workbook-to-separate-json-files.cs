using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchJsonExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing Excel files (XLS/XLSX)
            string inputFolder = @"C:\InputExcelFiles";

            // Output folder where JSON files will be saved
            string outputFolder = @"C:\OutputJsonFiles";

            // Verify input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all Excel files with .xls or .xlsx extensions
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.xls*");

            foreach (string excelPath in excelFiles)
            {
                try
                {
                    // Verify the source file still exists
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found (skipped): {excelPath}");
                        continue;
                    }

                    // Build the JSON file name based on the Excel file name
                    string jsonFileName = Path.GetFileNameWithoutExtension(excelPath) + ".json";
                    string jsonPath = Path.Combine(outputFolder, jsonFileName);

                    // Convert the Excel workbook to JSON using Aspose.Cells ConversionUtility
                    ConversionUtility.Convert(excelPath, jsonPath);

                    Console.WriteLine($"Converted: {excelPath} -> {jsonPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}