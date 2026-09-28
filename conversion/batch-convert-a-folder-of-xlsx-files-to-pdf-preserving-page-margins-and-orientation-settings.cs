// Title: Batch convert a folder of XLSX workbooks to PDF with Aspose.Cells while preserving page margins and orientation (C#)
// AI Prompts: Write a C# method that scans a given directory for *.xlsx files, uses Aspose.Cells.Utility.ConversionUtility to export each workbook to PDF, and keeps the original PageSetup (margins, orientation). | Generate C# code that creates the output folder if it does not exist, converts every Excel file in the source folder to PDF in bulk, and logs any conversion errors without aborting the batch.
// Common Searches: c# Aspose.Cells convert multiple xlsx files to pdf preserving page setup | how to keep Excel margins and orientation when exporting to PDF using Aspose.Cells | batch processing Excel to PDF with Aspose.Cells Utility in .NET | convert all workbooks in a folder to PDF with Aspose.Cells and handle errors per file | Aspose.Cells ConversionUtility example for folder-wide xlsx to pdf conversion
// Tags: Aspose.Cells batch XLSX to PDF conversion | retain page margins and orientation in PDF export | C# folder scan for Excel files conversion | ConversionUtility usage for bulk PDF generation | per‑file error logging in Aspose.Cells batch process

using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    // The sample enumerates every .xlsx file in a source directory, ensures the destination folder exists, and uses Aspose.Cells.Utility.ConversionUtility to convert each workbook to PDF while preserving its PageSetup settings such as margins and orientation. Conversion failures are caught per file, logged to the console, and the batch continues processing remaining files.
    public class Program
    {
        /// <summary>
        /// Entry point.
        /// </summary>
        public static void Main(string[] args)
        {
            // Example usage:
            //   args[0] = source folder containing XLSX files
            //   args[1] = destination folder for generated PDFs
            // If arguments are not supplied, use default demo folders.
            string sourceFolder = args.Length > 0 ? args[0] : @"C:\InputXlsx";
            string destFolder   = args.Length > 1 ? args[1] : @"C:\OutputPdf";

            try
            {
                BatchConvertXlsxToPdf(sourceFolder, destFolder);
                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }

        /// <summary>
        /// Converts every *.xlsx file found in <paramref name="sourceFolder"/>
        /// to a PDF file placed in <paramref name="destFolder"/>.
        /// </summary>
        /// <param name="sourceFolder">Folder containing source XLSX files.</param>
        /// <param name="destFolder">Folder where PDF files will be saved.</param>
        private static void BatchConvertXlsxToPdf(string sourceFolder, string destFolder)
        {
            // Validate source folder
            if (!Directory.Exists(sourceFolder))
                throw new DirectoryNotFoundException($"Source folder not found: {sourceFolder}");

            // Ensure destination folder exists
            Directory.CreateDirectory(destFolder);

            // Retrieve all XLSX files (case‑insensitive)
            string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string xlsxPath in xlsxFiles)
            {
                // Build output PDF path with same file name but .pdf extension
                string pdfFileName = Path.GetFileNameWithoutExtension(xlsxPath) + ".pdf";
                string pdfPath = Path.Combine(destFolder, pdfFileName);

                try
                {
                    // Perform conversion using the provided ConversionUtility rule.
                    // This method loads the workbook, respects its PageSetup (margins,
                    // orientation, etc.) and saves it as PDF.
                    ConversionUtility.Convert(xlsxPath, pdfPath);
                    Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {pdfFileName}");
                }
                catch (Exception fileEx)
                {
                    // Log the error but continue processing remaining files
                    Console.WriteLine($"Error converting '{xlsxPath}': {fileEx.Message}");
                }
            }
        }
    }
}
