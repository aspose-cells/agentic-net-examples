// Title: Convert multiple XLSX files to PDF/A‑1a in parallel with Aspose.Cells and C#
// Common Searches: c# convert all xlsx files in a folder to PDF/A-1a using Aspose.Cells | parallel processing of Excel to PDF/A conversion with Aspose.Cells in .NET | how to save an Aspose.Cells workbook as PDF/A-1a in C# | batch convert Excel workbooks to PDF/A-1a using multithreading in C# | Aspose.Cells example for PDF/A-1a compliance across many files
// Tags: Aspose.Cells batch XLSX to PDF/A conversion | PdfSaveOptions PdfA1a compliance C# | Directory enumeration Excel files Aspose.Cells | Multithreaded workbook save Aspose.Cells | C# parallel file conversion Aspose.Cells

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample enumerates all .xlsx files in a specified input directory, loads each workbook with Aspose.Cells, and saves it as a PDF/A‑1a document using PdfSaveOptions with Compliance set to PdfA1a. Files are processed concurrently with Parallel.ForEach, and the resulting PDFs are written to an output folder.
class Program
{
    static void Main(string[] args)
    {
        // Input directory containing XLSX files
        string inputDir = @"C:\InputXlsx";
        // Output directory for PDF/A‑1a files
        string outputDir = @"C:\OutputPdf";

        // Verify input directory exists
        if (!Directory.Exists(inputDir))
        {
            Console.Error.WriteLine($"Input directory does not exist: {inputDir}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputDir);

        // Get all XLSX files in the input directory (non‑recursive)
        string[] xlsxFiles = Directory.GetFiles(inputDir, "*.xlsx", SearchOption.TopDirectoryOnly);

        // Process files in parallel
        Parallel.ForEach(xlsxFiles, xlsxPath =>
        {
            try
            {
                // Verify the file still exists before loading
                if (!File.Exists(xlsxPath))
                {
                    Console.Error.WriteLine($"File not found: {xlsxPath}");
                    return;
                }

                // Load the workbook
                Workbook wb = new Workbook(xlsxPath);

                // Configure PDF/A‑1a save options
                PdfSaveOptions saveOptions = new PdfSaveOptions
                {
                    // Set compliance to PDF/A‑1a
                    Compliance = PdfCompliance.PdfA1a
                };

                // Build output file path with .pdf extension
                string fileName = Path.GetFileNameWithoutExtension(xlsxPath);
                string pdfPath = Path.Combine(outputDir, fileName + ".pdf");

                // Save the workbook as PDF/A‑1a
                wb.Save(pdfPath, saveOptions);
            }
            catch (Exception ex)
            {
                // Log or handle errors per file
                Console.Error.WriteLine($"Error processing '{xlsxPath}': {ex.Message}");
            }
        });

        Console.WriteLine("Conversion completed.");
    }
}
