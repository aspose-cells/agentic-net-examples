// Title: Batch convert Excel workbooks with WordArt to PDF while preserving gradient fills using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that scans a folder for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, and saves it as a PDF using PdfSaveOptions to retain WordArt gradient effects. | Generate C# code that creates an output directory if missing, iterates over all Excel files in a source folder, converts each to PDF, and logs any conversion errors without aborting the batch.
// Common Searches: asp.net batch export Excel files containing WordArt to PDF preserving gradients | c# convert multiple .xlsx workbooks to PDF with Aspose.Cells and keep shape formatting | how to automate PDF conversion of Excel workbooks that have WordArt objects | Aspose.Cells PdfSaveOptions settings for gradient rendering in exported PDFs
// Tags: batch Excel to PDF conversion Aspose.Cells | preserve WordArt gradient Aspose.Cells PDF export | C# directory iteration workbook conversion | PdfSaveOptions gradient rendering | automated Excel PDF generation .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample enumerates .xls, .xlsx, and .xlsm files in a specified input folder, loads each workbook with Aspose.Cells, and saves it as a PDF using default PdfSaveOptions, which retain WordArt gradient rendering. PDFs are written to an output folder, and any file‑specific errors are logged without stopping the batch process.
class Program
{
    static void Main()
    {
        // Folder containing the source Excel files
        string inputFolder = @"C:\InputSpreadsheets";

        // Folder where the resulting PDFs will be saved
        string outputFolder = @"C:\OutputPDFs";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify the input directory exists; if not, inform the user and exit gracefully
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Retrieve all Excel files (xls, xlsx, xlsm) in the input folder
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string excelPath in excelFiles)
        {
            string ext = Path.GetExtension(excelPath).ToLowerInvariant();
            if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm")
                continue; // Skip non‑Excel files

            // Verify the file exists before attempting to load
            if (!File.Exists(excelPath))
                continue;

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Configure PDF save options (default options are sufficient for WordArt rendering)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Determine the output PDF file name
                string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the workbook as PDF using the configured options
                workbook.Save(pdfPath, pdfOptions);
            }
            catch (Exception ex)
            {
                // Log or handle errors for individual files without stopping the whole process
                Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
            }
        }
    }
}
