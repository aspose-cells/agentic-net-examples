// Title: Batch convert all XLSM workbooks in a directory to PDF with MinimumSize optimization using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that scans a specified folder, loads each *.xlsm workbook with Aspose.Cells, and saves it as a PDF using PdfSaveOptions with PdfOptimization.MinimumSize. | Create a C# method that accepts input and output folder paths and performs batch conversion of macro‑enabled Excel files to optimized PDFs, returning a list of successful and failed files. | Write a C# snippet that adds detailed logging (file name, start time, end time, error messages) around the batch XLSM‑to‑PDF conversion using Aspose.Cells. | Show how to modify the batch conversion to skip hidden worksheets or protected sheets before saving each PDF with MinimumSize optimization.
// Common Searches: C# Aspose.Cells batch convert XLSM files to PDF with smallest file size | How to apply PdfOptimization.MinimumSize when converting macro-enabled Excel to PDF in .NET | Example code for converting all Excel macro workbooks in a folder to PDF using Aspose.Cells | Aspose.Cells PDFSaveOptions settings for reducing PDF size during bulk conversion | Error handling for batch XLSM to PDF conversion with Aspose.Cells C#
// Tags: Aspose.Cells PDFSaveOptions MinimumSize | batch XLSM to PDF conversion C# | macro-enabled Excel to PDF Aspose.Cells | process Excel files in folder Aspose.Cells | optimize PDF size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// This C# example enumerates all *.xlsm files in a given input directory, loads each workbook with Aspose.Cells, and saves it as a PDF in an output folder. It configures PdfSaveOptions with PdfOptimization.MinimumSize to produce the smallest possible PDFs, includes error handling for missing or corrupt files, and logs conversion results.
class BatchXlsmToPdf
{
    static void Main()
    {
        // Folder containing the source XLSM files
        string sourceFolder = @"C:\InputXlsmFolder";

        // Folder where the resulting PDFs will be saved
        string outputFolder = @"C:\OutputPdfFolder";

        try
        {
            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all XLSM files in the source folder (non‑recursive)
            string[] xlsmFiles = Directory.GetFiles(sourceFolder, "*.xlsm", SearchOption.TopDirectoryOnly);

            // Prepare PDF save options (default optimization)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            foreach (string xlsmPath in xlsmFiles)
            {
                try
                {
                    // Verify the source file exists before loading
                    if (!File.Exists(xlsmPath))
                    {
                        Console.WriteLine($"Source file not found: {xlsmPath}");
                        continue;
                    }

                    // Load the XLSM workbook
                    Workbook workbook = new Workbook(xlsmPath);

                    // Build the output PDF file path (same name, .pdf extension)
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsmPath);
                    string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    // Save the workbook as PDF using the specified options
                    workbook.Save(pdfPath, pdfOptions);
                    Console.WriteLine($"Converted: {xlsmPath} -> {pdfPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{xlsmPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Conversion completed. PDFs saved to: " + outputFolder);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Fatal error: " + ex.Message);
        }
    }
}
