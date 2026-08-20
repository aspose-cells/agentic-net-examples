// Title: Batch convert XLSM workbooks to Minimum‑Size PDFs with Aspose.Cells for .NET
// Description: A C# console utility that scans a source folder for *.xlsm files, converts each workbook to PDF using Aspose.Cells ConversionUtility, and applies PdfSaveOptions.OptimizationType = MinimumSize for compact output. The program creates matching PDFs in a target directory, logs progress, and isolates per‑file errors.
// Keywords: Aspose.Cells batch conversion | XLSM to PDF .NET | MinimumSize PDF optimization | ConversionUtility Convert | PdfSaveOptions OptimizationType | C# Excel macro PDF export | Windows folder processing | Aspose.Cells PDF compression
// Common Searches: how to batch convert xlsm files to pdf using aspose.cells | asp.net convert macro enabled excel to small pdf | c# convert all excel files in a folder to pdf with minimum size | aspocells pdfsaveoptions minimumsize example | automate folder conversion of xlsm to pdf
// Developer Intent: Convert every macro‑enabled Excel workbook in a directory to an optimized PDF using Aspose.Cells.
// Use Cases: Archive nightly macro‑driven financial reports as space‑efficient PDFs. | Provide a command‑line service that transforms uploaded XLSM files into web‑ready PDFs. | Integrate into a server‑side workflow that reduces storage costs by compressing Excel workbooks to MinimumSize PDFs.
// AI Prompts: Generate a C# method that receives input and output folder paths and batch converts all .xlsm files to PDF with MinimumSize optimization, including detailed error handling. | Write unit tests that verify a PDF is produced for each XLSM file and that PdfSaveOptions.OptimizationType is set to MinimumSize. | Create a PowerShell wrapper that runs the compiled converter, passes folder arguments, and writes a log file with success/failure for each conversion.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

namespace BatchXlsmToPdf
{
    // A C# console utility that scans a source folder for *.xlsm files, converts each workbook to PDF using Aspose.Cells ConversionUtility, and applies PdfSaveOptions.OptimizationType = MinimumSize for compact output. The program creates matching PDFs in a target directory, logs progress, and isolates per‑file errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the XLSM files
            string sourceFolder = @"C:\InputXlsm";
            // Folder where the PDF files will be saved
            string outputFolder = @"C:\OutputPdf";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Load options – let Aspose.Cells auto‑detect the format
            LoadOptions loadOptions = new LoadOptions();

            // PDF save options with MinimumSize optimization
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                OptimizationType = PdfOptimizationType.MinimumSize
            };

            try
            {
                // Process each .xlsm file in the source folder
                foreach (string xlsmPath in Directory.GetFiles(sourceFolder, "*.xlsm"))
                {
                    try
                    {
                        // Determine the output PDF file path
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsmPath);
                        string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                        // Convert the workbook to PDF
                        ConversionUtility.Convert(xlsmPath, loadOptions, pdfPath, pdfSaveOptions);

                        Console.WriteLine($"Converted: {xlsmPath} -> {pdfPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting '{xlsmPath}': {ex.Message}");
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
