// Title: How to enable PdfSaveOptions.CrossString for precise text placement when converting Excel to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets PdfSaveOptions.CrossString with specific X and Y coordinates to position a string in the PDF output of an Excel workbook. | Show a .NET example that configures Aspose.Cells PdfSaveOptions to place custom text at exact coordinates during Excel‑to‑PDF conversion. | Provide a step‑by‑step snippet that activates CrossString in PdfSaveOptions and saves a workbook as PDF with controlled text layout.
// Common Searches: Aspose.Cells CrossString setting for PDF conversion C# example | How to control text coordinates in PDF generated from Excel with Aspose.Cells .NET | Custom text placement using PDF save options when exporting workbook to PDF in C# | Precise PDF layout adjustments with Aspose.Cells rendering options
// Tags: CrossString property usage | custom text coordinates PDF Aspose.Cells | Excel to PDF layout precision C# | Aspose.Cells PDF rendering coordinate control | C# Aspose.Cells PDF conversion options

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, configures PdfSaveOptions (including enabling CrossString with custom X/Y coordinates for a string), and saves the workbook as a PDF while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (customize as needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Example: set one page per sheet (optional)
            pdfOptions.OnePagePerSheet = true;

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
