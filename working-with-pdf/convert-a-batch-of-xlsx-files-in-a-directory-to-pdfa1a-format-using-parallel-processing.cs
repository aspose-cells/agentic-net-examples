// Title: Batch convert XLSX files to PDF/A‑1a with Aspose.Cells using Parallel.ForEach in C#
// AI Prompts: Write C# code that scans a folder for .xlsx files and converts each workbook to PDF/A‑1a using Aspose.Cells inside a Parallel.ForEach loop. | Show how to configure Aspose.Cells PdfSaveOptions for PDF/A‑1a compliance and apply those options while processing multiple Excel files concurrently. | Create a C# routine that logs conversion errors and continues processing when converting a batch of Excel workbooks to PDF/A‑1a in parallel.
// Common Searches: C# Aspose.Cells convert all Excel files in a directory to PDF/A‑1a concurrently | How to use Parallel.ForEach with Aspose.Cells to batch export XLSX to PDF/A‑1a | Set PdfSaveOptions.Compliance to PdfA1a in a multithreaded Excel to PDF conversion | Batch processing of XLSX to PDF/A‑1a with error handling in .NET
// Tags: Aspose.Cells parallel batch XLSX to PDF/A conversion | PdfSaveOptions PDF/A‑1a compliance C# | Concurrent Excel workbook conversion .NET | Error handling during bulk PDF/A export Aspose | Folder based XLSX to PDF/A processing Aspose.Cells

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample enumerates all .xlsx files in a specified input folder, then uses Parallel.ForEach to load each workbook with Aspose.Cells, configures PdfSaveOptions for PDF/A‑1a compliance, and saves the resulting PDF to an output directory. Errors for individual files are logged without interrupting the overall batch conversion.
class Program
{
    static void Main(string[] args)
    {
        // Input and output directories (adjust as needed)
        string inputDirectory = @"C:\InputXlsx";
        string outputDirectory = @"C:\OutputPdfA";

        // Verify input directory exists
        if (!Directory.Exists(inputDirectory))
        {
            Console.Error.WriteLine($"Input directory not found: {inputDirectory}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputDirectory);

        // Get all .xlsx files in the input directory (top level only)
        string[] xlsxFiles = Directory.GetFiles(inputDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

        // Parallel conversion of each workbook to PDF/A-1a
        Parallel.ForEach(xlsxFiles, xlsxPath =>
        {
            try
            {
                // Verify the source file exists before attempting to load
                if (!File.Exists(xlsxPath))
                {
                    Console.Error.WriteLine($"Source file not found: {xlsxPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(xlsxPath);

                // Prepare PDF/A-1a save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Compliance = PdfCompliance.PdfA1a,
                    OnePagePerSheet = false
                };

                // Determine output PDF file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                string pdfPath = Path.Combine(outputDirectory, fileNameWithoutExt + ".pdf");

                // Save as PDF/A-1a using the options overload
                workbook.Save(pdfPath, pdfOptions);
            }
            catch (Exception ex)
            {
                // Log errors for the current file without stopping the batch
                Console.Error.WriteLine($"Error converting '{xlsxPath}': {ex.Message}");
            }
        });

        Console.WriteLine("Batch conversion completed.");
    }
}
