// Title: C# Batch Convert Excel Workbooks to PDF with OnePagePerSheet Using Aspose.Cells
// Description: A console utility that scans a folder for .xls, .xlsx, .xlsm, and .xlsb files, loads each workbook with Aspose.Cells, applies PdfSaveOptions.OnePagePerSheet, and saves a paginated PDF to a target directory while handling missing files and conversion errors.
// Keywords: Aspose.Cells | C# Excel to PDF | batch PDF conversion | OnePagePerSheet | process multiple Excel files | PdfSaveOptions | automate Excel PDF export
// Common Searches: convert all Excel files in a folder to PDF C# Aspose.Cells | Aspose.Cells OnePagePerSheet example | batch Excel to PDF conversion .NET | C# script to export workbooks as PDF | automate Excel PDF generation with Aspose
// Developer Intent: Automatically transform every Excel workbook in a specified directory into a PDF where each worksheet occupies a single page.
// Use Cases: Nightly generation of printable PDF reports from a repository of financial spreadsheets. | Command‑line tool for users to upload Excel files and receive paginated PDFs for consistent printing. | Server‑side watcher that converts newly added Excel files to PDFs with fixed pagination for archiving.
// AI Prompts: Write a reusable C# method that takes source and destination folder paths and converts all Excel files to PDF with OnePagePerSheet enabled using Aspose.Cells. | Explain how to extend the batch converter to recurse into subfolders and generate a CSV log of conversion results. | Show code to add support for password‑protected Excel workbooks during batch PDF conversion with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToPdfBatch
{
    // A console utility that scans a folder for .xls, .xlsx, .xlsm, and .xlsb files, loads each workbook with Aspose.Cells, applies PdfSaveOptions.OnePagePerSheet, and saves a paginated PDF to a target directory while handling missing files and conversion errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the folder containing Excel workbooks
            string sourceFolder = @"C:\InputExcelFiles";

            // Path to the folder where PDFs will be saved
            string outputFolder = @"C:\OutputPdfFiles";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder '{sourceFolder}' does not exist.");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            try
            {
                // Get all Excel files in the source folder (supports .xls, .xlsx, .xlsm, etc.)
                string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
                foreach (string excelPath in excelFiles)
                {
                    // Filter only Excel formats based on extension
                    string ext = Path.GetExtension(excelPath).ToLowerInvariant();
                    if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm" && ext != ".xlsb")
                        continue;

                    // Ensure the file still exists before loading
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found: '{excelPath}'. Skipping.");
                        continue;
                    }

                    try
                    {
                        // Load the workbook from the file
                        Workbook workbook = new Workbook(excelPath);

                        // Configure PDF save options to force one page per sheet
                        PdfSaveOptions pdfOptions = new PdfSaveOptions
                        {
                            OnePagePerSheet = true
                        };

                        // Build the output PDF file name (same base name as the Excel file)
                        string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                        string pdfPath = Path.Combine(outputFolder, pdfFileName);

                        // Save the workbook as PDF using the configured options
                        workbook.Save(pdfPath, pdfOptions);

                        Console.WriteLine($"Converted '{excelPath}' to PDF successfully.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting '{excelPath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
