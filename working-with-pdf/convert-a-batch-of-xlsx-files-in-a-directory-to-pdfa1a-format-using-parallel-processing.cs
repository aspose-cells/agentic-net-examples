// Title: Convert multiple XLSX files to PDF/A‑1a in C# with Aspose.Cells using Parallel.ForEach
// Common Searches: how to convert a folder of xlsx files to pdf/a-1a using Aspose.Cells C# | parallel processing of Excel to PDF/A conversion with Aspose.Cells | batch convert Excel workbooks to PDF/A-1a in .NET | Aspose.Cells example for saving workbook as PDF/A-1a | C# code to convert all .xlsx in a directory to PDF/A-1a
// Tags: Aspose.Cells parallel processing | PDF/A-1a conversion with PdfSaveOptions | C# batch Excel to PDF conversion | directory based XLSX to PDF/A conversion | PdfSaveOptions Compliance PdfA1a

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample enumerates every .xlsx file in a specified input folder, creates an output folder, and uses Parallel.ForEach to convert each workbook concurrently. For each file it loads the workbook with Aspose.Cells, sets PdfSaveOptions.Compliance to PdfA1a for PDF/A‑1a compliance, and saves the result as a .pdf file in the output directory, while handling per‑file and global exceptions.
class BatchXlsxToPdfA
{
    static void Main(string[] args)
    {
        try
        {
            // Input directory containing XLSX files
            string inputDir = @"C:\InputXlsx";
            // Output directory for PDF/A‑1a files
            string outputDir = @"C:\OutputPdfA";

            // Ensure output directory exists
            Directory.CreateDirectory(outputDir);

            // Get all .xlsx files in the input directory (non‑recursive)
            string[] xlsxFiles = Directory.GetFiles(inputDir, "*.xlsx", SearchOption.TopDirectoryOnly);

            // Process files in parallel
            Parallel.ForEach(xlsxFiles, xlsxPath =>
            {
                try
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(xlsxPath))
                    {
                        Console.Error.WriteLine($"File not found: '{xlsxPath}'");
                        return;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(xlsxPath);

                    // Configure PDF save options for PDF/A‑1a compliance
                    PdfSaveOptions saveOptions = new PdfSaveOptions
                    {
                        Compliance = PdfCompliance.PdfA1a
                        // EmbedStandardFonts property is not available in this version
                    };

                    // Build output file path with same name but .pdf extension
                    string outputFileName = Path.GetFileNameWithoutExtension(xlsxPath) + ".pdf";
                    string outputPath = Path.Combine(outputDir, outputFileName);

                    // Save the workbook as PDF/A‑1a
                    workbook.Save(outputPath, saveOptions);
                }
                catch (Exception ex)
                {
                    // Log or handle errors per file as needed
                    Console.Error.WriteLine($"Error processing '{xlsxPath}': {ex.Message}");
                }
            });

            Console.WriteLine("Batch conversion completed.");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.Error.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
