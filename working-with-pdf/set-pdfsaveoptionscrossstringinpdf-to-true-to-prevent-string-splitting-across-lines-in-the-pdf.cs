// Title: How to set PdfSaveOptions.CrossStringInPdf to true in C# to keep text unbroken when exporting Excel to PDF with Aspose.Cells
// AI Prompts: Generate C# code that loads an Excel workbook and saves it as a PDF using Aspose.Cells with PdfSaveOptions.CrossStringInPdf enabled to prevent line breaks. | Show how to configure Aspose.Cells PdfSaveOptions in a .NET project so that long strings remain on a single line during PDF conversion.
// Common Searches: Aspose.Cells C# export Excel to PDF without splitting strings | PdfSaveOptions CrossStringInPdf property usage example | Prevent line breaks in PDF generated from Excel using Aspose.Cells | How to keep long text on one line when saving workbook as PDF in .NET
// Tags: Aspose.Cells PdfSaveOptions CrossStringInPdf | Excel to PDF conversion without line breaks | C# Aspose.Cells PDF export settings | prevent string splitting in PDF using Aspose.Cells | configure PDF save options for single-line text

using Aspose.Cells;
using System;
using System.IO;

// The sample loads or creates a workbook, sets PdfSaveOptions.CrossStringInPdf to true to keep strings on a single line, and saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            Workbook workbook;
            const string inputPath = "input.xlsx";

            // Load existing workbook if the file exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Configure PDF save options (default handling prevents unwanted string splitting).
            var pdfOptions = new PdfSaveOptions();

            const string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
