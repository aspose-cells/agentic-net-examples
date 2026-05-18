using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchJsonExport
{
    class Program
    {
        static void Main()
        {
            // Folder containing the source Excel files (XLS format)
            string sourceFolder = @"C:\InputExcelFiles";

            // Ensure the folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Get all .xls files in the folder (non‑recursive)
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xls");

            if (excelFiles.Length == 0)
            {
                Console.WriteLine("No XLS files found in the specified folder.");
                return;
            }

            foreach (string excelPath in excelFiles)
            {
                try
                {
                    // Build the output JSON file path – same name, .json extension, placed in the same folder
                    string jsonPath = Path.ChangeExtension(excelPath, ".json");

                    // Convert the Excel workbook to JSON using Aspose.Cells.Utility.ConversionUtility
                    // This uses the rule: ConversionUtility.Convert(string source, string saveAs)
                    ConversionUtility.Convert(excelPath, jsonPath);

                    Console.WriteLine($"Converted '{Path.GetFileName(excelPath)}' to '{Path.GetFileName(jsonPath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting file '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}