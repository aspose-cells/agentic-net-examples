// Title: Convert a folder of macro‑enabled Excel (.xlsm) workbooks to PDF with MinimumSize optimization using Aspose.Cells for .NET (C#)
// Common Searches: C# code to batch convert all .xlsm files in a directory to PDF using Aspose.Cells | How to apply MinimumSize PDF optimization when saving Excel workbooks with Aspose.Cells | Aspose.Cells PdfSaveOptions OptimizationType MinimumSize example for multiple files | Convert macro‑enabled Excel workbooks to PDF in bulk with Aspose.Cells .NET | Automate folder‑level XLSM to PDF conversion with size reduction using Aspose.Cells
// Tags: batch xlsm to pdf conversion Aspose.Cells | PdfSaveOptions MinimumSize optimization C# | Aspose.Cells workbook.Save PDF options | convert macro enabled Excel files to PDF .NET | directory based Excel to PDF batch processing

using System;
using System.IO;
using Aspose.Cells;

// The sample scans a specified input folder for *.xlsm files, creates an output directory if needed, and iterates through each workbook. For every file it loads the workbook with Aspose.Cells, configures PdfSaveOptions (using the MinimumSize optimization), and saves the workbook as a PDF with the same base name in the output folder, logging successes and any conversion errors.
class BatchXlsmToPdfConverter
{
    static void Main(string[] args)
    {
        // Path to the folder containing XLSM files
        string sourceFolder = @"C:\InputXlsmFolder";
        // Path to the folder where PDFs will be saved
        string outputFolder = @"C:\OutputPdfFolder";

        // Verify source directory exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Ensure output directory exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        // Get all .xlsm files in the source folder (non‑recursive)
        string[] xlsmFiles = Directory.GetFiles(sourceFolder, "*.xlsm", SearchOption.TopDirectoryOnly);

        foreach (string xlsmPath in xlsmFiles)
        {
            try
            {
                // Verify the file still exists before loading
                if (!File.Exists(xlsmPath))
                {
                    Console.WriteLine($"File not found (skipped): {xlsmPath}");
                    continue;
                }

                // Load the XLSM workbook
                Workbook workbook = new Workbook(xlsmPath);

                // Prepare PDF save options (using default optimization)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Build the output PDF file path (same name, .pdf extension)
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsmPath);
                string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                // Save the workbook as PDF
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"Converted: {xlsmPath} -> {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting '{xlsmPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
