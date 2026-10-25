// Title: Convert all Excel files in a folder to PDF with one page per sheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells batch convert Excel files in a directory to PDF with one page per sheet | How to export multiple .xlsx workbooks to PDF using Aspose.Cells while preserving sheet pagination | Aspose.Cells PdfSaveOptions OnePagePerSheet example for processing a folder of workbooks in .NET | Automate conversion of .xls, .xlsx, .xlsm, and .xlsb files to PDF in C# | Save each worksheet as a separate PDF page using Aspose.Cells C#
// Tags: batch Excel to PDF conversion Aspose.Cells | PdfSaveOptions pagination control | folder enumeration of Excel workbooks C# | convert .xls .xlsx .xlsm .xlsb to PDF | non‑recursive directory scan C#

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToPdfBatch
{
    // The program scans a specified input folder for .xls, .xlsx, .xlsm, and .xlsb files, loads each workbook with Aspose.Cells, and saves it as a PDF in an output folder using PdfSaveOptions configured with OnePagePerSheet = true, ensuring each worksheet appears on a single PDF page.
    class Program
    {
        static void Main(string[] args)
        {
            // Define source folder containing Excel files and destination folder for PDFs
            string sourceFolder = @"C:\InputExcelFiles";
            string outputFolder = @"C:\OutputPdfFiles";

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Perform the conversion
            ConvertFolderToPdf(sourceFolder, outputFolder);

            Console.WriteLine("Conversion completed.");
        }

        /// <param name="sourceFolder">Folder containing Excel files.</param>
        /// <param name="outputFolder">Folder where PDF files will be saved.</param>
        static void ConvertFolderToPdf(string sourceFolder, string outputFolder)
        {
            // Define supported Excel extensions
            string[] extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.xlsb" };

            // Gather all Excel files from the source folder (non‑recursive)
            var excelFiles = new System.Collections.Generic.List<string>();
            foreach (var ext in extensions)
            {
                excelFiles.AddRange(Directory.GetFiles(sourceFolder, ext, SearchOption.TopDirectoryOnly));
            }

            // Configure PDF save options to force one page per sheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Process each Excel file
            foreach (string excelPath in excelFiles)
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(excelPath);

                // Build the output PDF file name (same base name, .pdf extension)
                string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the workbook as PDF with the specified options
                workbook.Save(pdfPath, pdfOptions);
            }
        }
    }
}
