// Title: Convert all XLSM files in a directory to PDF with MinimumSize optimization using Aspose.Cells for .NET
// AI Prompts: Generate a C# console application that scans a folder for *.xlsm files and saves each workbook as a PDF using Aspose.Cells with PdfSaveOptions set to MinimumSize. | Write code to batch process macro‑enabled Excel workbooks, applying the MinimumSize PDF optimization when exporting to PDF via Aspose.Cells. | Create a script that reads XLSM files from an input directory, converts them to compressed PDFs using PdfSaveOptions.OptimizationType = MinimumSize, and writes the PDFs to an output folder.
// Common Searches: aspnet batch convert xlsm to pdf minimum size Aspose.Cells example | c# program to export macro enabled Excel files to compressed PDF using Aspose | how to use PdfSaveOptions OptimizationType MinimumSize for multiple workbooks in .NET
// Tags: batch xlsm to pdf conversion Aspose.Cells | PdfSaveOptions MinimumSize optimization | export macro-enabled workbook to pdf .NET | C# directory file enumeration Aspose.Cells | convert multiple Excel workbooks to compressed PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// A C# console program that enumerates all *.xlsm files in a specified input folder, loads each workbook with Aspose.Cells, and saves it as a PDF using PdfSaveOptions with OptimizationType set to MinimumSize, writing the resulting PDFs to an output directory while handling errors gracefully.
class BatchXlsmToPdfConverter
{
    static void Main()
    {
        // Folder containing the source XLSM files
        string sourceFolder = @"C:\InputFolder";

        // Folder where the resulting PDF files will be saved
        string outputFolder = @"C:\OutputFolder";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Get all XLSM files in the source folder
        string[] xlsmFiles = Directory.GetFiles(sourceFolder, "*.xlsm", SearchOption.TopDirectoryOnly);

        foreach (string xlsmPath in xlsmFiles)
        {
            try
            {
                // Ensure the file exists before loading
                if (!File.Exists(xlsmPath))
                {
                    Console.WriteLine($"File not found: {xlsmPath}");
                    continue;
                }

                // Load the XLSM workbook
                Workbook workbook = new Workbook(xlsmPath);

                // Configure PDF save options to use MinimumSize optimization
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OptimizationType = PdfOptimizationType.MinimumSize
                };

                // Build the output PDF file path
                string pdfFileName = Path.GetFileNameWithoutExtension(xlsmPath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the workbook as PDF with the specified options
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"Converted: {Path.GetFileName(xlsmPath)} -> {pdfFileName}");
            }
            catch (Exception ex)
            {
                // Log conversion errors without stopping the batch process
                Console.WriteLine($"Error converting '{xlsmPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Conversion completed. PDFs are saved in: " + outputFolder);
    }
}
