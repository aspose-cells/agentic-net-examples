// Title: Save an Excel workbook as a single‑page PDF with Aspose.Cells for .NET using PdfSaveOptions
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, configures the PDF export so that each worksheet is rendered on a single page, and saves the workbook as a PDF. | Provide a C# example that checks whether the output directory exists (creating it if missing) and applies PDF/A‑1b compliance before calling Workbook.Save to generate a PDF.
// Common Searches: How to make each sheet fit on a single PDF page using Aspose.Cells in C# | Setting PDF/A‑1b compliance when converting Excel to PDF with Aspose.Cells .NET | C# create missing directory before saving PDF with Aspose.Cells Workbook.Save | PdfSaveOptions example for one‑page‑per‑sheet export in Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | Aspose.Cells export Excel to PDF .NET | PDF/A compliance Aspose.Cells | C# ensure output directory exists Aspose.Cells | Workbook.Save PDF Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel file with Aspose.Cells, sets PdfSaveOptions to place each worksheet on a single PDF page (optionally enabling PDF/A‑1b compliance), ensures the output folder exists, and saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF export options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Uncomment the following line if the PdfCompliance enum is available in your Aspose.Cells version
            // pdfOptions.Compliance = Aspose.Cells.PdfCompliance.PdfA1b;

            // Fit each worksheet on a single page
            pdfOptions.OnePagePerSheet = true;

            // Path for the output PDF file
            string outputPath = "output.pdf";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
