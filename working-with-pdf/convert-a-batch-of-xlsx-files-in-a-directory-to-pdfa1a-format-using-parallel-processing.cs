// Title: Parallel Batch Convert XLSX to PDF/A‑1a with Aspose.Cells for .NET (C#)
// Description: Scans a folder for .xlsx files and uses Aspose.Cells with Parallel.ForEach to convert each workbook to PDF/A‑1a on all CPU cores, reporting success or errors in the console.
// Keywords: Aspose.Cells C# | XLSX to PDF/A-1a conversion | parallel Excel PDF conversion | batch PDF/A-1a generation | PdfSaveOptions Compliance PdfA1a | LoadOptions Xlsx | ConversionUtility Aspose | .NET multi‑core processing | archive‑ready PDF | Windows console Excel conversion
// Common Searches: convert folder of xlsx files to pdf/a-1a c# | asp.net parallel excel to pdf/a batch conversion | aspocells batch pdf/a-1a example | c# multi‑threaded xlsx to pdf/a conversion | how to use ConversionUtility for pdf/a compliance
// Developer Intent: Convert every XLSX file in a directory to PDF/A‑1a using multi‑core parallelism.
// Use Cases: Automated archival of financial spreadsheets on a server farm. | Real‑time compliance processing of incoming Excel reports in a shared drop folder. | Integration into CI/CD pipelines to verify PDF/A‑1a output of generated workbooks.
// AI Prompts: Write C# code that adds a CancellationToken to the Parallel.ForEach loop for graceful shutdown. | Show how to log conversion progress and errors to a rotating file while preserving parallel execution. | Extend the program to recursively process sub‑folders and preserve the original directory structure in the output.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

namespace BatchXlsxToPdfA
{
    // Scans a folder for .xlsx files and uses Aspose.Cells with Parallel.ForEach to convert each workbook to PDF/A‑1a on all CPU cores, reporting success or errors in the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing the XLSX files. Change as needed.
            string sourceDirectory = @"C:\InputXlsx";

            // Verify the directory exists.
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory does not exist: {sourceDirectory}");
                return;
            }

            // Get all .xlsx files in the directory (non‑recursive).
            string[] xlsxFiles = Directory.GetFiles(sourceDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

            if (xlsxFiles.Length == 0)
            {
                Console.WriteLine("No XLSX files found to convert.");
                return;
            }

            // Process files in parallel, using the number of logical processors.
            ParallelOptions parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = Environment.ProcessorCount
            };

            Parallel.ForEach(xlsxFiles, parallelOptions, xlsxPath =>
            {
                try
                {
                    // Destination PDF file path (same name, .pdf extension).
                    string pdfPath = Path.ChangeExtension(xlsxPath, ".pdf");

                    // Load options for XLSX format.
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                    // Save options for PDF/A‑1a compliance.
                    PdfSaveOptions saveOptions = new PdfSaveOptions
                    {
                        Compliance = PdfCompliance.PdfA1a
                    };

                    // Perform the conversion using the provided ConversionUtility method.
                    ConversionUtility.Convert(xlsxPath, loadOptions, pdfPath, saveOptions);

                    Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} → {Path.GetFileName(pdfPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{Path.GetFileName(xlsxPath)}': {ex.Message}");
                }
            });

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
