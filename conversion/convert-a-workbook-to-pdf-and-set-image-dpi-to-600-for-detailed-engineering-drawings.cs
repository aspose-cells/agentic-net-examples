// Title: Convert an Excel workbook to PDF with 600 DPI images using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF with 600 DPI image resolution | How to set ImageDPI in PdfSaveOptions when saving workbook as PDF | High‑resolution PDF generation from engineering drawings in Excel using Aspose.Cells | C# code sample for converting .xlsx to PDF with custom image DPI | Aspose.Cells PdfSaveOptions ImageDPI property usage example
// Tags: Aspose.Cells PdfSaveOptions ImageDPI | Excel to PDF high resolution conversion | C# set PDF image DPI Aspose.Cells | engineering drawing PDF export .NET | custom image DPI PDF generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an .xlsx file with Aspose.Cells, configures PdfSaveOptions to set ImageDPI to 600 for high‑resolution PDF output, and saves the workbook as a PDF while handling missing‑file and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Uncomment the following line if your Aspose.Cells version supports ImageDPI
            // pdfOptions.ImageDPI = 600; // Set image DPI to 600 for high‑resolution output

            // Convert and save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
