// Title: Convert an Excel workbook to PDF using Aspose.Cells for .NET and enable Fast Web View for faster browser rendering
// Common Searches: how to export .xlsx to PDF with fast web view using Aspose.Cells C# | Aspose.Cells PDFSaveOptions fast web view setting example | C# generate PDF from Excel optimized for web view with Aspose.Cells | fast web view not available in Aspose.Cells PDF export workaround | save Excel workbook as PDF for quick loading in browsers Aspose.Cells .NET
// Tags: Aspose.Cells PDFSaveOptions configuration | Excel to PDF conversion .NET | Fast Web View PDF generation Aspose.Cells | browser-friendly PDF export C# | handling missing Fast Web View support Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the source Excel file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (noting that Fast Web View is not supported in the current version), and saves the workbook as a PDF, providing console feedback and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (Fast Web View is not supported directly in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
