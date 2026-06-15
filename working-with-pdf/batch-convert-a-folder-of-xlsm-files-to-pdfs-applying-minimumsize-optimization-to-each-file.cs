using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

namespace BatchXlsmToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the XLSM files
            string sourceFolder = @"C:\InputXlsm";
            // Folder where the PDF files will be saved
            string outputFolder = @"C:\OutputPdf";

            try
            {
                // Verify source folder exists
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder not found: {sourceFolder}");
                    return;
                }

                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Get all .xlsm files in the source folder (non‑recursive)
                string[] xlsmFiles = Directory.GetFiles(sourceFolder, "*.xlsm");

                foreach (string xlsmPath in xlsmFiles)
                {
                    try
                    {
                        // Verify the source file exists (redundant but safe)
                        if (!File.Exists(xlsmPath))
                        {
                            Console.WriteLine($"Source file not found: {xlsmPath}");
                            continue;
                        }

                        // Build the corresponding PDF file path
                        string pdfFileName = Path.GetFileNameWithoutExtension(xlsmPath) + ".pdf";
                        string pdfPath = Path.Combine(outputFolder, pdfFileName);

                        // Load options – let Aspose.Cells auto‑detect the format
                        LoadOptions loadOptions = new LoadOptions();

                        // PDF save options with MinimumSize optimization
                        PdfSaveOptions pdfOptions = new PdfSaveOptions
                        {
                            OptimizationType = PdfOptimizationType.MinimumSize
                        };

                        // Convert the XLSM file to PDF using the utility method
                        ConversionUtility.Convert(xlsmPath, loadOptions, pdfPath, pdfOptions);

                        Console.WriteLine($"Converted: {xlsmPath} -> {pdfPath}");
                    }
                    catch (Exception ex)
                    {
                        // Log any conversion errors but continue processing other files
                        Console.WriteLine($"Error converting '{xlsmPath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}