// Title: Convert an Excel workbook to PDF without Fast Web View using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook as PDF without fast web view | How to generate a single continuous PDF stream from an Excel file with Aspose.Cells | Disable FastWebView in PDFSaveOptions when converting Excel to PDF in .NET | Create a non‑linearized PDF from .xlsx using Aspose.Cells | Aspose.Cells PDF conversion options for continuous output
// Tags: Aspose.Cells PDFSaveOptions usage | Excel workbook PDF conversion C# | single continuous PDF output Aspose.Cells | disable linearized PDF Aspose.Cells | non‑fastwebview PDF generation .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an .xlsx file into an Aspose.Cells Workbook, creates a PdfSaveOptions object (which does not expose a FastWebView setting), and saves the workbook as a PDF. The resulting PDF is a single continuous stream without Fast Web View (linearized) features.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (no FastWebView property in Aspose.Cells)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
