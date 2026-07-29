// Title: C# Batch Convert XLSM Workbooks to Optimized PDF (Minimum Size) with Aspose.Cells
// Description: A console utility that scans a folder for *.xlsm files, creates an output directory, and uses Aspose.Cells ConversionUtility with LoadOptions and PdfSaveOptions (OptimizationType = MinimumSize) to generate matching PDF files. Includes basic error handling and progress logging.
// Keywords: Aspose.Cells | C# | .NET | batch convert XLSM to PDF | PDF MinimumSize optimization | ConversionUtility | LoadOptions | PdfSaveOptions | macro-enabled Excel conversion | folder processing
// Common Searches: convert all xlsm files in a folder to pdf using aspose.cells | aspnet batch conversion of macro enabled workbooks to optimized pdf | c# code for pdf minimum size optimization with aspose.cells | aspse.cells convert multiple excel files to pdf command line | how to use ConversionUtility to batch convert xlsm to pdf
// Developer Intent: Automatically transform every XLSM workbook in a specified directory into a PDF with MinimumSize compression.
// Use Cases: Generate lightweight PDFs from a collection of macro‑enabled reports for archiving. | Schedule nightly export of XLSM dashboards to PDF for stakeholder distribution. | Create a command‑line tool that processes a folder of XLSM files and saves size‑optimized PDFs to another folder.
// AI Prompts: Write C# code that uses Aspose.Cells to batch convert all XLSM files in a directory to PDF with MinimumSize optimization and logs each conversion. | Show how to configure LoadOptions for opening macro‑enabled workbooks and set PdfSaveOptions for MinimumSize in Aspose.Cells. | Explain how to extend the sample to recursively process subfolders while preserving the original folder hierarchy in the output location.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

namespace BatchXlsmToPdf
{
    // A console utility that scans a folder for *.xlsm files, creates an output directory, and uses Aspose.Cells ConversionUtility with LoadOptions and PdfSaveOptions (OptimizationType = MinimumSize) to generate matching PDF files. Includes basic error handling and progress logging.
    class Program
    {
        static void Main()
        {
            try
            {
                // Folder containing the source XLSM files
                string sourceFolder = @"C:\InputXlsm";

                // Folder where the converted PDF files will be saved
                string outputFolder = @"C:\OutputPdf";

                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Get all XLSM files in the source folder
                string[] xlsmFiles = Directory.GetFiles(sourceFolder, "*.xlsm", SearchOption.TopDirectoryOnly);

                // Load options for XLSM files (use Xlsx format as LoadFormat.Xlsm is not defined)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // PDF save options with MinimumSize optimization
                PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
                {
                    OptimizationType = PdfOptimizationType.MinimumSize
                };

                // Convert each XLSM file to PDF using ConversionUtility
                foreach (string sourcePath in xlsmFiles)
                {
                    if (!File.Exists(sourcePath))
                    {
                        Console.WriteLine($"Source file not found: {sourcePath}");
                        continue;
                    }

                    try
                    {
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                        string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                        // Perform conversion
                        ConversionUtility.Convert(sourcePath, loadOptions, destPath, pdfSaveOptions);

                        Console.WriteLine($"Converted: {sourcePath} -> {destPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting file '{sourcePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
