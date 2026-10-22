// Title: Convert an Excel workbook to PDF and attempt to enable full‑screen view with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert XLSX to PDF with full screen view | How to set PDF to open in full screen when exporting Excel using Aspose.Cells .NET | Aspose.Cells PDFSaveOptions open in full screen mode not working | Save Excel workbook as PDF and launch in presentation mode with Aspose.Cells C# | Aspose.Cells PDF export full screen option missing in .NET SDK
// Tags: Aspose.Cells PDFSaveOptions full-screen | Excel to PDF conversion C# Aspose.Cells | enable PDF full-screen view Aspose.Cells | Workbook.Save PDF Aspose.Cells .NET | Aspose.Cells PDF export presentation mode

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing .xlsx workbook, creates a PdfSaveOptions instance, and saves the workbook as a PDF using Aspose.Cells for .NET. It highlights that the OpenInFullScreenMode property is unavailable in the current SDK, so the PDF cannot be forced to open in full‑screen mode directly.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (OpenInFullScreenMode not available in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
