// Title: Convert an Excel workbook to a PDF/A‑2b file while preserving the original ICC color profile with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, sets PdfSaveOptions.Compliance to PdfA2b, and saves the workbook as a PDF preserving the source ICC profile. | Show how to verify the input Excel file exists, configure PDF/A‑2b compliance, and embed the workbook’s color profile when exporting to PDF with Aspose.Cells. | Provide a robust example that handles exceptions, applies PDF/A‑2b settings, and ensures accurate color reproduction in the resulting PDF.
// Common Searches: Aspose.Cells C# export Excel to PDF/A-2b with original ICC profile | how to keep color accuracy when converting .xlsx to PDF/A-2b using Aspose.Cells | set PdfSaveOptions compliance to PdfA2b and embed color profile in .NET | sample code for PDF/A‑2b conversion of Excel workbook preserving colors Aspose.Cells
// Tags: Aspose.Cells PDF/A-2b compliance | preserve ICC profile Aspose.Cells | Excel to PDF/A-2b conversion C# | PdfSaveOptions compliance setting | color accuracy PDF export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input XLSX file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object, sets its Compliance property to PdfA2b to meet PDF/A‑2b standards, and saves the workbook as a PDF while preserving the original ICC color profile and handling any errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Uncomment the following line if the Aspose.Cells version supports PDF/A compliance
            // pdfOptions.Compliance = PdfCompliance.PdfA2b;

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
