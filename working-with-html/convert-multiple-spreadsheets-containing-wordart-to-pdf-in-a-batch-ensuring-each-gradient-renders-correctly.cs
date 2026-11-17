// Title: Batch convert Excel workbooks with WordArt to PDF in C# using Aspose.Cells while preserving gradient fills
// Common Searches: C# Aspose.Cells batch convert multiple Excel files to PDF preserving WordArt gradients | How to export .xls/.xlsx workbooks containing WordArt to PDF with correct gradient rendering | Save Excel worksheets with WordArt as PDF in a folder using Aspose.Cells | Automate conversion of Excel files to PDF while keeping shape gradients in C# | Aspose.Cells PdfSaveOptions default settings keep WordArt gradient colors
// Tags: batch Excel to PDF conversion Aspose.Cells | preserve WordArt gradient in PDF Aspose.Cells | C# process multiple workbooks to PDF | PdfSaveOptions for gradient rendering | convert .xls .xlsx .xlsm files to PDF in bulk

using System;
using System.IO;
using Aspose.Cells;

namespace BatchWordArtToPdfApp
{
    // The C# program scans a given directory for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, and saves it as a PDF using default PdfSaveOptions, which retain WordArt gradient rendering, writing the resulting PDFs to a specified output folder.
    class BatchWordArtToPdf
    {
        static void Main()
        {
            // Folder containing the Excel files with WordArt
            string inputFolder = @"C:\InputExcelFiles";
            // Folder where the resulting PDFs will be saved
            string outputFolder = @"C:\OutputPdfFiles";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            string[] excelFiles;
            try
            {
                // Get all Excel files in the input folder (supports .xls, .xlsx, .xlsm)
                excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error accessing input folder: {ex.Message}");
                return;
            }

            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                    continue; // Skip non‑Excel files

                // Verify the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook inside a using block for proper disposal
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        // Configure PDF save options (default handling preserves WordArt gradients)
                        PdfSaveOptions pdfOptions = new PdfSaveOptions();

                        // Build output PDF file name
                        string pdfFileName = Path.GetFileNameWithoutExtension(filePath) + ".pdf";
                        string pdfPath = Path.Combine(outputFolder, pdfFileName);

                        // Save the workbook as PDF
                        workbook.Save(pdfPath, pdfOptions);

                        Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to PDF successfully.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
