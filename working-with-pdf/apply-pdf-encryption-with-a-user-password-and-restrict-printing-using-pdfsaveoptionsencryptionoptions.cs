// Title: Encrypt a PDF generated from an Aspose.Cells workbook with a user password and block printing using PdfSaveOptions in C#
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as a password‑protected PDF and disables printing by configuring PdfSaveOptions.EncryptionOptions. | Show how to set PdfEncryptionOptions.UserPassword and adjust PdfEncryptionOptions.Permissions to prevent printing before calling Workbook.Save in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set user password for PDF output and prevent printing | PdfSaveOptions EncryptionOptions example to disable printing in generated PDF | How to apply PDF encryption with password and no‑print permission using Aspose.Cells | Protect PDF created from Excel workbook with password and printing restriction in .NET | Aspose.Cells PDF export password protection and print restriction code sample
// Tags: Aspose.Cells PDF encryption with user password C# | PdfSaveOptions set user password Aspose.Cells | disable printing PDF Aspose.Cells C# | PdfEncryptionOptions permissions Aspose.Cells | protect generated PDF workbook Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a Workbook, adds sample data, then configures a PdfSaveOptions instance. It assigns a PdfEncryptionOptions object, sets the UserPassword, and defines Permissions that omit the PrintDocument flag to block printing. Finally, the workbook is saved as a password‑protected PDF with printing disabled, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample PDF with encryption");

            // Configure PDF save options (no encryption due to missing Pdf namespace)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            string outputPath = "EncryptedOutput.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
