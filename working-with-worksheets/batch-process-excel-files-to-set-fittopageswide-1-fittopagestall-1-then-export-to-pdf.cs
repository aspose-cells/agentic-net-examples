// Title: Batch Convert Excel Files to Single‑Page PDF with Aspose.Cells (.NET)
// Description: A C# console app that scans a folder for Excel workbooks (xls, xlsx, xlsm, csv), sets each worksheet's PageSetup.FitToPagesWide and FitToPagesTall to 1 for a one‑page layout, and saves the files as PDFs using Aspose.Cells PdfSaveOptions. The program creates the output directory, logs progress, and handles supported formats only.
// Keywords: Aspose.Cells | C# | batch Excel to PDF | FitToPagesWide | FitToPagesTall | single page PDF | page setup scaling | PdfSaveOptions | automate Excel conversion | .NET PDF export | command line Excel PDF
// Common Searches: how to batch convert Excel to PDF with Aspose.Cells | set FitToPagesWide and FitToPagesTall for all worksheets in C# | export multiple Excel files as single‑page PDFs | Aspose.Cells PDFSaveOptions example for folder processing | C# code to convert a directory of spreadsheets to PDF
// Developer Intent: Automatically convert every Excel workbook in a specified directory to a PDF where each worksheet fits on one page.
// Use Cases: Produce printable PDFs of monthly financial statements, ensuring each sheet prints on a single page. | Archive CSV imports as uniformly formatted PDFs for compliance documentation. | Run a nightly job that transforms newly uploaded Excel reports into PDF for distribution to non‑Excel users.
// AI Prompts: Generate C# code that uses Aspose.Cells to batch convert Excel files in a folder to PDF with FitToPagesWide=1 and FitToPagesTall=1. | Show how to add custom PDF metadata (author, title, subject) to the batch conversion while keeping the one‑page scaling. | Provide recommendations for adding error handling, logging, and progress reporting to the Excel‑to‑PDF loop.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace BatchExcelToPdf
{
    // A C# console app that scans a folder for Excel workbooks (xls, xlsx, xlsm, csv), sets each worksheet's PageSetup.FitToPagesWide and FitToPagesTall to 1 for a one‑page layout, and saves the files as PDFs using Aspose.Cells PdfSaveOptions. The program creates the output directory, logs progress, and handles supported formats only.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing Excel files
            string inputFolder = @"C:\InputExcel";
            // Folder to save generated PDFs
            string outputFolder = @"C:\OutputPdf";

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files (xls, xlsx, csv, etc.) in the input folder
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                // Filter supported Excel formats
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm" && ext != ".csv")
                    continue;

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Set page setup for each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.PageSetup.FitToPagesWide = 1;   // Fit to 1 page wide
                    sheet.PageSetup.FitToPagesTall = 1;   // Fit to 1 page tall
                }

                // Prepare PDF save options (optional customizations can be added here)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Build output PDF file path
                string pdfFileName = Path.GetFileNameWithoutExtension(filePath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the workbook as PDF
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to PDF successfully.");
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
