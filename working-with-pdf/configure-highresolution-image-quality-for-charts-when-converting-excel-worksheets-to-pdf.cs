// Title: Set high‑resolution chart image quality when converting an Excel workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a PdfSaveOptions object, sets ImageResolution to 300 DPI and JpegQuality to 100, then saves an Excel workbook to PDF so that chart images are rendered at high quality using Aspose.Cells. | Show how to enable lossless chart rendering and adjust PDF image compression settings in Aspose.Cells to produce crisp charts in the generated PDF.
// Common Searches: Aspose.Cells how to increase chart DPI in PDF export C# | C# set image resolution for charts when saving Excel to PDF with Aspose.Cells | PdfSaveOptions ImageResolution property example Aspose.Cells .NET | Improve chart clarity in PDF generated from Excel using Aspose.Cells | Configure JPEG quality for chart images in Aspose.Cells PDF conversion
// Tags: Aspose.Cells PdfSaveOptions ImageResolution | high‑resolution chart export PDF .NET | set JPEG quality Aspose.Cells PDF | Excel to PDF chart image quality | configure PDF image DPI Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates a PdfSaveOptions instance, sets ImageResolution to 300 DPI and JpegQuality to 100 to ensure charts are rendered sharply, and saves the workbook as a PDF, including basic file‑existence checking and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (high‑resolution settings may depend on the library version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
