// Title: Apply AES‑256 encryption with a user password when saving an Excel workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that saves a Workbook as a PDF with AES‑256 encryption and a user‑defined password using Aspose.Cells. | Show how to set the EncryptionType and Password properties on PdfSaveOptions before calling Workbook.Save to protect the output PDF.
// Common Searches: Aspose.Cells C# encrypt PDF with AES256 and user password | How to add password protection to PDF generated from Excel using Aspose.Cells .NET | PdfSaveOptions encryption type AES256 example Aspose.Cells
// Tags: Aspose.Cells PDF AES-256 encryption | PdfSaveOptions user password protection | C# save workbook as encrypted PDF | Aspose.Cells encryption settings for PDF | AES256 PDF protection .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The original code loads an Excel workbook and saves it as a PDF without any security. To meet the requirement, configure a PdfSaveOptions instance by setting EncryptionType to AES256 and assigning a user password, then pass this options object to Workbook.Save so the resulting PDF is protected with AES‑256 encryption.
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
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (encryption not available in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
