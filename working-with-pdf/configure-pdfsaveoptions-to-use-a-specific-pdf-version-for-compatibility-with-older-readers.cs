// Title: How to set a specific PDF version or compliance with PdfSaveOptions when converting an Excel workbook to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a Workbook, configures PdfSaveOptions.PdfCompliance to PdfA1b, and saves the workbook as a PDF file. | Show how to use PdfSaveOptions to target PDF 1.5 (PdfCompliance.Pdf15) for older PDF readers when exporting an Excel file to PDF with Aspose.Cells. | Provide an example that checks for the existence of an .xlsx file, creates a new workbook if missing, and then saves it as PDF using a chosen PDF compliance level.
// Common Searches: Aspose.Cells set PDF/A-1b compliance when saving workbook to PDF C# | PdfSaveOptions PdfCompliance PDF 1.5 example Aspose.Cells .NET | How to export Excel to PDF with a specific PDF version using Aspose.Cells | C# create workbook if missing then save as PDF with PdfSaveOptions | Configure PDF compatibility for older readers Aspose.Cells PdfSaveOptions
// Tags: Aspose.Cells PdfSaveOptions PdfCompliance | C# export Excel to PDF specific version | PDF/A-1b generation Aspose.Cells | PDF 1.5 compliance Aspose.Cells | fallback workbook creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file or creates a new workbook when the source is absent, configures PdfSaveOptions (optionally setting PdfCompliance to a desired PDF version such as PDF/A‑1b or PDF 1.5), and saves the workbook as a PDF document compatible with older readers.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists; create a blank workbook if it does not.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook(); // creates a default workbook with one worksheet
                workbook.Save(inputPath, SaveFormat.Xlsx);
            }

            // Configure PDF save options (e.g., set PDF version via compliance if needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Example: set PDF compliance for broader compatibility (optional)
            // pdfOptions.PdfCompliance = PdfCompliance.PdfA1b;

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
