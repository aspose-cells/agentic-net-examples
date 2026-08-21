// Title: Batch Convert XLSX Files to PDF with Page Margins & Orientation Using Aspose.Cells for .NET
// Description: C# program that validates a source directory, creates an output folder, enumerates all *.xlsx files, and uses Aspose.Cells.Utility.ConversionUtility.Convert to generate PDF files. The conversion respects each workbook's page setup (margins, orientation) and logs success or error for every file.
// Keywords: Aspose.Cells | C# | .NET | batch conversion | XLSX to PDF | preserve margins | preserve orientation | ConversionUtility | folder conversion | Excel to PDF automation | page setup | command line tool
// Common Searches: batch convert xlsx to pdf asp.net | aspocells preserve page orientation pdf | convert folder of excel files to pdf c# | aspocells ConversionUtility example | keep margins when converting excel to pdf
// Developer Intent: Convert every XLSX workbook in a given folder to PDF while retaining the original page margins and orientation.
// Use Cases: Automate nightly generation of PDF reports from a directory of Excel files without losing layout. | Provide a lightweight command‑line utility for users to transform uploaded Excel templates into PDFs that keep original formatting. | Integrate the batch converter into a server‑side archiving process that stores Excel documents as PDFs for compliance, preserving exact page setup.
// AI Prompts: Generate C# code that extends this batch converter to also handle .xls and .xlsm files while preserving page setup. | Write a PowerShell script that calls XlsxToPdfBatchConverter with user‑supplied source and output paths and captures any errors. | Explain how to modify the program to write conversion results (file name, status, timestamp) to a CSV log instead of the console.

using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    // C# program that validates a source directory, creates an output folder, enumerates all *.xlsx files, and uses Aspose.Cells.Utility.ConversionUtility.Convert to generate PDF files. The conversion respects each workbook's page setup (margins, orientation) and logs success or error for every file.
    public class XlsxToPdfBatchConverter
    {
        /// <param name="sourceFolder">Folder containing the source .xlsx files.</param>
        /// <param name="outputFolder">Folder where the resulting PDF files will be saved.</param>
        public static void Run(string sourceFolder, string outputFolder)
        {
            // Validate source folder
            if (!Directory.Exists(sourceFolder))
            {
                throw new DirectoryNotFoundException($"Source folder not found: {sourceFolder}");
            }

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Get all .xlsx files (including .xlsm, .xltx, etc. if needed)
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string excelPath in excelFiles)
            {
                // Verify the file still exists before processing
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found (skipped): {excelPath}");
                    continue;
                }

                // Build PDF file name with same base name
                string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                try
                {
                    // Perform conversion using Aspose.Cells.Utility.ConversionUtility
                    // This method respects the workbook's page setup (margins, orientation, etc.)
                    ConversionUtility.Convert(excelPath, pdfPath);
                    Console.WriteLine($"Converted: {excelPath} -> {pdfPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }

    internal class Program
    {
        /// <summary>
        /// Entry point of the application.
        /// </summary>
        /// <param name="args">
        /// args[0] - source folder path (optional)
        /// args[1] - output folder path (optional)
        /// </param>
        private static void Main(string[] args)
        {
            try
            {
                string sourceFolder;
                string outputFolder;

                if (args.Length >= 2)
                {
                    sourceFolder = args[0];
                    outputFolder = args[1];
                }
                else
                {
                    // Default example paths – adjust as needed
                    sourceFolder = @"C:\InputXlsx";
                    outputFolder = @"C:\OutputPdf";
                }

                XlsxToPdfBatchConverter.Run(sourceFolder, outputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
