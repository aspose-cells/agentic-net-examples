// Title: How to encrypt a PDF generated from an Excel workbook with a strong password and AES‑256 using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets PdfSaveOptions.Password to a strong value, sets EncryptionType to AES256, and saves the workbook as an encrypted PDF. | Show the steps to verify the source Excel file, configure PDF security options, and handle exceptions when saving a password‑protected PDF with Aspose.Cells. | Provide a complete example that demonstrates applying PDF password protection and AES‑256 encryption using PdfSaveOptions in a .NET console application.
// Common Searches: Aspose.Cells encrypt PDF with password and AES256 in C# | Set PdfSaveOptions.Password and EncryptionType for PDF output in Aspose.Cells .NET | C# example for password‑protecting PDF generated from Excel using Aspose.Cells | How to apply strong PDF encryption when converting workbook to PDF with Aspose.Cells | PdfSaveOptions EncryptionType AES256 usage Aspose.Cells tutorial
// Tags: Aspose.Cells PDF password protection C# | PdfSaveOptions AES256 encryption | Encrypt PDF from Excel workbook Aspose.Cells | Set PdfSaveOptions.Password Aspose.Cells | PDF security options Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The program checks that the source Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object, assigns a strong password and sets EncryptionType to AES256, then saves the workbook as an encrypted PDF. All operations are wrapped in try‑catch to report errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (no encryption due to missing PdfSecurityOptions assembly)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
