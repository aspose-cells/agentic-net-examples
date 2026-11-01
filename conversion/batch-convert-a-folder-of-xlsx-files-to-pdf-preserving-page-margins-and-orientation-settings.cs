// Title: Batch convert XLSX files in a folder to PDF with original page margins and orientation using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells convert all Excel files in a directory to PDF preserving layout | How to export multiple .xlsx workbooks to PDF without changing page setup using Aspose.Cells | Batch PDF conversion of Excel files while keeping margins and orientation in .NET | Aspose.Cells PdfSaveOptions OnePagePerSheet false example for folder conversion | Automate folder‑wide XLSX to PDF conversion with Aspose.Cells in a console app
// Tags: Aspose.Cells folder Excel to PDF conversion | PdfSaveOptions OnePagePerSheet false usage | preserve worksheet page layout during PDF export | C# console batch conversion of XLSX to PDF | maintain margins and orientation Aspose.Cells PDF

using System;
using System.IO;
using Aspose.Cells;

// The console application scans a specified input directory for .xlsx files, loads each workbook with Aspose.Cells, configures PdfSaveOptions (OnePagePerSheet = false) to keep the original worksheet layout, and saves each workbook as a PDF in an output folder, creating the folder if needed and handling missing files or conversion errors.
class BatchXlsxToPdfConverter
{
    static void Main(string[] args)
    {
        // Input folder containing XLSX files
        string inputFolder = @"C:\InputXlsxFolder";
        // Output folder where PDF files will be saved
        string outputFolder = @"C:\OutputPdfFolder";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure output directory exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        // Get all XLSX files in the input folder (non‑recursive)
        string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string xlsxPath in xlsxFiles)
        {
            try
            {
                // Verify the file still exists before loading
                if (!File.Exists(xlsxPath))
                {
                    Console.WriteLine($"File not found: {xlsxPath}");
                    continue;
                }

                // Load the workbook from the XLSX file
                Workbook workbook = new Workbook(xlsxPath);

                // Create PDF save options; keep default compliance
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Ensure that the page layout defined in each worksheet is used
                    OnePagePerSheet = false
                    // Compliance property omitted for compatibility with all versions
                };

                // Build the output PDF file name (same base name, .pdf extension)
                string pdfFileName = Path.GetFileNameWithoutExtension(xlsxPath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the workbook as PDF using the defined options
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {pdfFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting file '{xlsxPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
