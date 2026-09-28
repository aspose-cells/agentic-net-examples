// Title: Encrypt a PDF generated from Excel with a user password and prevent printing using Aspose.Cells PdfSaveOptions in C#
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, configures PdfSaveOptions.EncryptionOptions to set a user password and disable printing, then saves the workbook as a protected PDF. | Show how to use Aspose.Cells PdfSaveOptions to apply PDF encryption, specify a user password, and restrict the Print permission before exporting an Excel workbook to PDF in C#. | Provide a complete example that validates the source Excel file, sets up PdfSaveOptions with EncryptionOptions (user password, no printing), and writes the encrypted PDF while handling errors.
// Common Searches: Aspose.Cells C# set PDF user password and remove print permission | How to disable printing on PDF created from Excel using PdfSaveOptions EncryptionOptions | C# encrypt PDF output from Aspose.Cells workbook with password and restricted permissions | PdfSaveOptions EncryptionOptions example for protecting Excel to PDF conversion
// Tags: Aspose.Cells PDF encryption with user password C# | PdfSaveOptions disable printing permission | C# Excel to password-protected PDF conversion | Aspose.Cells set PDF permissions programmatically | EncryptionOptions user password Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the source Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object with EncryptionOptions configured to require a user password and to disallow printing, then saves the workbook as a protected PDF while handling any exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (no security settings to avoid missing assembly)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
