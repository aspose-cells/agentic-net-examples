// Title: C# – Batch Convert XLSX Files to PDF with Aspose.Cells while Keeping Margins & Orientation
// Description: A console utility that validates a folder path, enumerates all *.xlsx files, and uses Aspose.Cells.Utility.ConversionUtility.Convert to create same‑named PDFs. The conversion respects each workbook’s page‑setup settings (margins and orientation). It logs each success, reports per‑file errors, and prints a final count of processed files.
// Keywords: Aspose.Cells | C# batch XLSX to PDF | Excel to PDF conversion .NET | preserve margins orientation | ConversionUtility | folder conversion | command line Excel PDF | automate Excel PDF export
// Common Searches: convert all excel files in a folder to pdf c# | aspocells batch xlsx to pdf conversion | preserve page margins when exporting Excel to PDF | command line tool to convert folder of xlsx to pdf | c# console app batch convert excel to pdf
// Developer Intent: Convert every .xlsx workbook in a specified directory to PDF while retaining the original page layout (margins and orientation) using Aspose.Cells.
// Use Cases: Nightly automation that turns a batch of financial Excel reports into PDF for archiving. | A lightweight console tool that end‑users run to generate PDFs from a folder of spreadsheets without opening Excel. | Integration into a larger ETL pipeline where incoming Excel files are instantly converted to PDF for downstream processing. | Providing a self‑service utility for legal teams to preserve exact worksheet formatting when sharing PDFs.
// AI Prompts: Write C# code that uses Aspose.Cells to batch convert all .xlsx files in a directory to PDF, ensuring original margins and orientation are kept. | Add detailed error handling and a log file to the batch conversion loop, capturing file names and exception messages for any failures. | Show how to configure ConversionUtility options to set PDF page size, embed fonts, or adjust image quality while processing multiple Excel files.

using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsExample
{
    // A console utility that validates a folder path, enumerates all *.xlsx files, and uses Aspose.Cells.Utility.ConversionUtility.Convert to create same‑named PDFs. The conversion respects each workbook’s page‑setup settings (margins and orientation). It logs each success, reports per‑file errors, and prints a final count of processed files.
    public class XlsxToPdfBatchConverter
    {
        /// <param name="folderPath">Full path of the folder containing XLSX files.</param>
        public static void ConvertFolder(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
                throw new ArgumentException("Folder path must be provided.", nameof(folderPath));

            if (!Directory.Exists(folderPath))
                throw new DirectoryNotFoundException($"The folder '{folderPath}' does not exist.");

            // Get all .xlsx files (case‑insensitive) in the folder
            string[] xlsxFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string sourceFile in xlsxFiles)
            {
                try
                {
                    if (!File.Exists(sourceFile))
                    {
                        Console.WriteLine($"Source file not found: {sourceFile}");
                        continue;
                    }

                    // Build the destination PDF file name in the same folder
                    string destFile = Path.ChangeExtension(sourceFile, ".pdf");

                    // Perform the conversion.
                    ConversionUtility.Convert(sourceFile, destFile);

                    Console.WriteLine($"Converted: {Path.GetFileName(sourceFile)} -> {Path.GetFileName(destFile)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{sourceFile}': {ex.Message}");
                }
            }

            Console.WriteLine($"Batch conversion completed. Processed {xlsxFiles.Length} file(s).");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                string folderPath;

                if (args.Length > 0)
                {
                    folderPath = args[0];
                }
                else
                {
                    Console.Write("Enter the full path of the folder containing XLSX files: ");
                    folderPath = Console.ReadLine();
                }

                XlsxToPdfBatchConverter.ConvertFolder(folderPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
