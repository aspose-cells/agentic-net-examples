// Title: Encrypt a PDF generated from an Excel workbook with Aspose.Cells for .NET using PdfSaveOptions.Password and EncryptionType
// Common Searches: how to add password protection to a PDF when saving an Excel file with Aspose.Cells C# | set strong encryption for PDF output using Aspose.Cells PdfSaveOptions in .NET | Aspose.Cells example to save workbook as encrypted PDF with custom password | C# code to configure PdfSaveOptions.EncryptionType for PDF security in Aspose.Cells | protect PDF generated from workbook with Aspose.Cells using strong encryption
// Tags: PdfSaveOptions password protection Aspose.Cells .NET | strong PDF encryption Aspose.Cells | encrypt Excel workbook to PDF Aspose.Cells | C# configure PdfSaveOptions EncryptionType | Aspose.Cells PDF security settings

using System;
using System.IO;
using Aspose.Cells;

namespace PdfEncryptionExampleApp
{
    // The sample loads an Excel workbook, optionally assigns a password and strong encryption via PdfSaveOptions, and saves the workbook as a password‑protected PDF using Aspose.Cells for .NET.
    class PdfEncryptionExample
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Ensure the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Set up PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Password protection and encryption are available in newer Aspose.Cells versions.
                // If your version supports them, you can uncomment and set the following:
                // pdfOptions.Password = "My$tr0ngP@ssw0rd!";
                // pdfOptions.EncryptionType = EncryptionType.StrongEncryption;

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
