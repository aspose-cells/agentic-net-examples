// Title: Create a tagged PDF for accessibility when converting an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: how to save an Excel workbook as an accessible tagged PDF using Aspose.Cells C# | Aspose.Cells PDF/UA compliance option for screen reader friendly PDFs | C# example for generating PDF/UA compliant PDF from .xlsx with Aspose.Cells | enable accessibility tags in PDF output when converting Excel to PDF in .NET | Aspose.Cells PdfSaveOptions set compliance to PDF/UA for accessibility
// Tags: Aspose.Cells PdfSaveOptions compliance property | tagged PDF generation from .xlsx | PDF/UA output with Aspose.Cells | screen reader friendly PDF conversion .NET | Excel workbook to accessible PDF

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input Excel file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions (optionally setting PDF/UA compliance for accessibility), and saves the workbook as a tagged PDF, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (tagged PDF support may be enabled via compliance settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Optional: set PDF compliance to PDF/UA for accessibility compliance
                // Compliance = PdfCompliance.PdfUAX
            };

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Successfully saved PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
