// Title: Convert an Excel workbook to PDF using Aspose.Cells for .NET while acknowledging that font embedding cannot be turned off in the current release
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells and saves it as a PDF, explicitly showing that PdfSaveOptions does not expose a font‑embedding toggle. | Demonstrate how to use Aspose.Cells PdfSaveOptions to create a PDF from a workbook and explain the impact on file size when font embedding cannot be disabled. | Provide a complete example with error handling that converts a workbook to PDF, logs the limitation of font‑embedding control, and suggests alternative size‑reduction techniques.
// Common Searches: Aspose.Cells .NET how to disable font embedding when saving PDF | reduce PDF file size when converting Excel to PDF with Aspose.Cells | PdfSaveOptions missing font embedding option Aspose.Cells version | convert .xlsx to PDF without embedding fonts using Aspose.Cells | Aspose.Cells PDF conversion size optimization tips
// Tags: Aspose.Cells PDF conversion options .NET | PdfSaveOptions font embedding limitation | Excel to PDF size reduction Aspose.Cells | disable font embedding Aspose.Cells PDF | workbook to PDF without embedded fonts

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the existence of input.xlsx, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions instance (which in this version does not provide a font‑embedding switch), and saves the workbook as output.pdf. Basic exception handling reports missing files or runtime errors, and the code comments clarify that font embedding cannot be disabled in the current Aspose.Cells release.
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

            // Create PDF save options (font embedding control is not available in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF file saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
