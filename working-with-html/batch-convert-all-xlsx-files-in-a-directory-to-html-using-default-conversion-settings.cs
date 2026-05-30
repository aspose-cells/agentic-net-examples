using System;
using System.IO;
using Aspose.Cells.Utility;

namespace BatchXlsxToHtml
{
    class Program
    {
        static void Main(string[] args)
        {
            // Specify the folder containing the XLSX files.
            // You can change this path as needed.
            string sourceFolder = @"C:\InputXlsx";

            // Verify that the folder exists.
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Get all .xlsx files in the folder (non‑recursive).
            string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string xlsxPath in xlsxFiles)
            {
                // Build the output HTML file path by changing the extension to .html.
                string htmlPath = Path.ChangeExtension(xlsxPath, ".html");

                try
                {
                    // Use Aspose.Cells ConversionUtility with default settings.
                    // This follows the provided rule: ConversionUtility.Convert(string, string)
                    ConversionUtility.Convert(xlsxPath, htmlPath);
                    Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {Path.GetFileName(htmlPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting {xlsxPath}: {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}