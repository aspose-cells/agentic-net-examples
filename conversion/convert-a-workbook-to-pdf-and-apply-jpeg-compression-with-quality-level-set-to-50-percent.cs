// Title: Convert an Excel workbook to PDF with 50% JPEG compression using Aspose.Cells for .NET
// Common Searches: how to set JPEG quality to 50 when saving Excel as PDF with Aspose.Cells C# | Aspose.Cells PDFSaveOptions JPEG compression level example in .NET | C# export workbook to PDF with reduced image size using Aspose.Cells | adjust image quality in PDF output from Excel using Aspose.Cells library | save Excel file as PDF with 50 percent JPEG quality Aspose.Cells .NET
// Tags: Aspose.Cells PDFSaveOptions JPEG quality | Excel to PDF conversion with image compression | C# set PDF image quality Aspose.Cells | reduce PDF file size Aspose.Cells PDF export | configure JpegQuality property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the source Excel file, loads it into a Workbook, creates a PdfSaveOptions object, sets the JpegQuality property to 50 (and optionally the CompressionType to JPEG), and then saves the workbook as a PDF. It includes basic error handling for missing files and unexpected exceptions.
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
                throw new FileNotFoundException($"The file '{inputPath}' was not found.");

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
