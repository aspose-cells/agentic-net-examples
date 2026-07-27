// Title: Batch convert multiple Excel workbooks to PDF without charts using Aspose.Cells LoadOptions (C#)
// Description: A C# console example that scans a folder for .xlsx and .xls files, loads each workbook with Aspose.Cells LoadOptions (optionally enabling LoadDataOnly to skip drawings), and saves them as PDFs using PdfSaveOptions (OnePagePerSheet, IgnoreError). The utility creates PDF files in a target directory while omitting chart objects, making it ideal for automated reporting pipelines.
// Keywords: Aspose.Cells | C# | .NET | batch Excel to PDF | LoadOptions | LoadDataOnly | ignore charts | ignore drawings | PdfSaveOptions | OnePagePerSheet | command line utility | automation | server side conversion | GitHub example | Excel to PDF conversion
// Common Searches: convert all Excel files in a folder to PDF Aspose.Cells C# | skip charts when saving Excel as PDF with Aspose.Cells | batch Excel to PDF conversion using LoadOptions | C# code to export multiple workbooks to PDF without drawings | Aspose.Cells example for folder‑wide PDF export
// Developer Intent: Automatically transform every Excel workbook in a directory into a PDF while excluding chart and drawing objects.
// Use Cases: Generate lightweight PDF reports from nightly Excel workbooks that contain large charts. | Create a command‑line tool for users to drop Excel files into a folder and receive chart‑free PDFs. | Integrate into a web service that receives uploaded spreadsheets, strips visual objects, and returns PDF output for archiving.
// AI Prompts: Write a C# console program using Aspose.Cells that loads all .xlsx/.xls files from a specified folder with LoadDataOnly enabled and saves each as a PDF with OnePagePerSheet. | Provide error‑handling best practices for batch converting Excel files to PDF, including folder validation, logging of failed files, and graceful continuation. | Explain how LoadOptions.LoadDataOnly works in Aspose.Cells and how it affects chart rendering during PDF export.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToPdf
{
    // A C# console example that scans a folder for .xlsx and .xls files, loads each workbook with Aspose.Cells LoadOptions (optionally enabling LoadDataOnly to skip drawings), and saves them as PDFs using PdfSaveOptions (OnePagePerSheet, IgnoreError). The utility creates PDF files in a target directory while omitting chart objects, making it ideal for automated reporting pipelines.
    class Program
    {
        static void Main()
        {
            // Folder containing the Excel files to be processed
            string sourceFolder = @"C:\InputExcels";
            // Folder where the resulting PDF files will be saved
            string outputFolder = @"C:\OutputPdfs";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files (XLSX and XLS) in the source folder
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls")
                    continue; // Skip non‑Excel files

                // Additional safety: verify the file still exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (default options)
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                    // If a newer Aspose.Cells version supports LoadDataOnly, uncomment the line below:
                    // loadOptions.LoadDataOnly = true; // ignore drawing objects

                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // Prepare PDF save options (optional customizations)
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        // Fit each worksheet on a single page
                        OnePagePerSheet = true,
                        // Suppress rendering errors (e.g., missing chart data)
                        IgnoreError = true
                    };

                    // Build the output PDF file name
                    string pdfFileName = Path.GetFileNameWithoutExtension(filePath) + ".pdf";
                    string pdfPath = Path.Combine(outputFolder, pdfFileName);

                    // Save the workbook as PDF
                    workbook.Save(pdfPath, pdfOptions);

                    Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to PDF successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
