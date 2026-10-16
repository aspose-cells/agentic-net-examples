// Title: Convert an Excel workbook to PDF with Aspose.Cells for .NET while preserving image quality
// Common Searches: how to export an .xlsx file to PDF using Aspose.Cells in C# without losing image quality | Aspose.Cells C# PDF export retain original graphics fidelity | C# save Excel workbook as PDF with lossless image compression using Aspose.Cells | using PdfSaveOptions to generate high‑quality PDF from Excel in .NET | Aspose.Cells PDF conversion example with default image settings
// Tags: Aspose.Cells PDF generation C# | PdfSaveOptions image fidelity | Excel to PDF conversion .NET | lossless image handling Aspose.Cells | preserve graphics quality PDF Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PdfSaveOptions

// The example loads an existing .xlsx workbook with Aspose.Cells, creates a PdfSaveOptions object using its default configuration (which keeps image data unchanged), and saves the workbook as a PDF file, including error handling for missing input files.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine(ex.Message);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
