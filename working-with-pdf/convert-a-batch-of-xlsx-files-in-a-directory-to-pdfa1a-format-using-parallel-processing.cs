// Title: Batch Convert XLSX to PDF/A‑1a in Parallel with Aspose.Cells for .NET (C#)
// Description: A C# console app that scans a folder for .xlsx files, creates an output directory, and uses Parallel.ForEach with Aspose.Cells' ConversionUtility, LoadOptions and PdfSaveOptions to generate PDF/A‑1a compliant PDFs. Includes per‑file error handling and progress logging.
// Keywords: Aspose.Cells | C# | .NET | XLSX to PDF/A-1a conversion | parallel batch conversion | ConversionUtility | PdfSaveOptions | PDF/A-1a compliance | Excel archival | GitHub example | open source code
// Common Searches: convert folder of xlsx to pdf/a-1a c# | aspose.cells parallel batch conversion example | c# code to generate PDF/A-1a from Excel | batch convert excel to pdf/a using asp.net | github aspose.cells pdfa1a sample
// Developer Intent: Convert every XLSX file in a directory to a PDF/A‑1a document using concurrent processing.
// Use Cases: Automate archival of daily financial workbooks by converting each spreadsheet to PDF/A‑1a in a background service. | Accelerate migration of a large Excel repository to PDF/A‑1a by processing files concurrently with Aspose.Cells. | Integrate the batch converter into CI/CD pipelines to validate PDF/A compliance of generated reports before release.
// AI Prompts: Show how to preserve original file timestamps after conversion. | Add timing metrics to log conversion duration per file and total elapsed time. | Modify the code to traverse subfolders recursively while keeping parallel execution.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

namespace BatchXlsxToPdfA1a
{
    // A C# console app that scans a folder for .xlsx files, creates an output directory, and uses Parallel.ForEach with Aspose.Cells' ConversionUtility, LoadOptions and PdfSaveOptions to generate PDF/A‑1a compliant PDFs. Includes per‑file error handling and progress logging.
    class Program
    {
        static void Main()
        {
            try
            {
                // Directory containing the source XLSX files
                string sourceDirectory = @"C:\InputXlsx";

                // Verify source directory exists
                if (!Directory.Exists(sourceDirectory))
                {
                    Console.WriteLine($"Source directory not found: {sourceDirectory}");
                    return;
                }

                // Directory where the converted PDF/A‑1a files will be saved
                string outputDirectory = @"C:\OutputPdfA1a";

                // Ensure the output directory exists
                Directory.CreateDirectory(outputDirectory);

                // Get all .xlsx files in the source directory (non‑recursive)
                string[] xlsxFiles = Directory.GetFiles(sourceDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

                // Process each file in parallel
                Parallel.ForEach(xlsxFiles, sourcePath =>
                {
                    try
                    {
                        // Verify the source file still exists
                        if (!File.Exists(sourcePath))
                        {
                            Console.WriteLine($"Source file not found: {sourcePath}");
                            return;
                        }

                        // Build the destination PDF file path (same file name, .pdf extension)
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                        string destPath = Path.Combine(outputDirectory, fileNameWithoutExt + ".pdf");

                        // Load options – explicitly specify the source format (XLSX)
                        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                        // Save options – set PDF compliance to PDF/A‑1a
                        PdfSaveOptions saveOptions = new PdfSaveOptions
                        {
                            Compliance = PdfCompliance.PdfA1a
                        };

                        // Perform the conversion using the utility method that accepts options
                        ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);

                        Console.WriteLine($"Converted: {sourcePath} -> {destPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                    }
                });

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
