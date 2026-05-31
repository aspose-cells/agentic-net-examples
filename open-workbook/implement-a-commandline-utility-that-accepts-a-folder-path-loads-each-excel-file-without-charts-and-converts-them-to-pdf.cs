using System;
using System.IO;
using Aspose.Cells.Utility;

namespace ExcelToPdfBatch
{
    class Program
    {
        // List of extensions that represent Excel workbook files
        private static readonly string[] ExcelExtensions = new[]
        {
            ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".csv", ".tsv", ".html", ".mhtml"
        };

        static void Main(string[] args)
        {
            // Expect a single argument: the folder containing Excel files
            if (args.Length != 1)
            {
                Console.WriteLine("Usage: ExcelToPdfBatch <folderPath>");
                return;
            }

            string folderPath = args[0];

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Error: Folder \"{folderPath}\" does not exist.");
                return;
            }

            // Process each file in the folder
            foreach (string filePath in Directory.GetFiles(folderPath))
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();

                // Skip non‑Excel files
                if (Array.IndexOf(ExcelExtensions, extension) < 0)
                    continue;

                string pdfPath = Path.ChangeExtension(filePath, ".pdf");

                try
                {
                    // Convert the Excel workbook to PDF using Aspose.Cells utility method
                    ConversionUtility.Convert(filePath, pdfPath);
                    Console.WriteLine($"Converted: \"{Path.GetFileName(filePath)}\" → \"{Path.GetFileName(pdfPath)}\"");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to convert \"{Path.GetFileName(filePath)}\": {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}