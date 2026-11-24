// Title: Convert a folder of Excel workbooks to PDF with one page per sheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells batch convert all Excel files in a directory to PDF with one page per worksheet | How to save multiple .xls and .xlsx workbooks as PDFs using PdfSaveOptions.OnePagePerSheet in .NET | Automate folder‑wide Excel to PDF conversion with Aspose.Cells and preserve single‑sheet pagination | Aspose.Cells code to export every workbook in a folder to PDF with one‑page‑per‑sheet layout | Convert Excel files to PDF in bulk using Aspose.Cells PDFSaveOptions in a console application
// Tags: Aspose.Cells batch Excel to PDF conversion | PdfSaveOptions OnePagePerSheet setting | C# directory processing of Excel workbooks | export multiple workbooks as PDF Aspose.Cells | automated Excel to PDF conversion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The console program scans a specified input folder for .xls and .xlsx files, creates the output folder if it does not exist, loads each workbook with Aspose.Cells, applies PdfSaveOptions with OnePagePerSheet set to true, and saves each workbook as a PDF using the same base name. It logs each conversion and handles any errors, enabling bulk conversion of Excel workbooks to PDFs with consistent one‑page‑per‑sheet pagination.
class ExcelToPdfConverter
{
    static void Main(string[] args)
    {
        // Input folder containing Excel workbooks
        string inputFolder = @"C:\InputExcelFiles";
        // Output folder for generated PDFs
        string outputFolder = @"C:\OutputPdfFiles";

        // Ensure the output directory exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        // Get all Excel files in the input folder (supports .xls and .xlsx)
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx")
                continue; // Skip non‑Excel files

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Configure PDF save options to force one page per worksheet
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                };

                // Determine output PDF file name
                string pdfFileName = Path.GetFileNameWithoutExtension(filePath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the workbook as PDF with the specified options
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"Converted: {Path.GetFileName(filePath)} -> {pdfFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Conversion process completed.");
    }
}
