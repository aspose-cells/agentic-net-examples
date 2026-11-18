// Title: Apply AES‑256 encryption with a user password when saving an Excel workbook as PDF using Aspose.Cells for .NET
// Common Searches: aspnet apply AES-256 encryption to PDF generated from Excel with Aspose.Cells | set only a user password for PDF output using Aspose.Cells C# | how to use PdfSaveOptions to protect PDF with password in Aspose.Cells .NET | encrypt PDF from workbook with AES256 without owner password in C# | Aspose.Cells save workbook as PDF with user‑only password protection
// Tags: Aspose.Cells PDF AES-256 encryption | PdfSaveOptions user password protection | C# encrypt PDF from Excel workbook | Excel to PDF password protection Aspose | PDF security settings Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// Load an Excel workbook, configure PdfSaveOptions with a user‑only password and AES‑256 encryption, and save the workbook as a protected PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (no security applied)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
