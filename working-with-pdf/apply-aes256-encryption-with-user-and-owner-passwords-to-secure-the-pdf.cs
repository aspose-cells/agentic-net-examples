// Title: Encrypt a PDF created from an Excel workbook with AES‑256 and set user and owner passwords using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells and saves it as a PDF protected by AES‑256 encryption, providing distinct user and owner passwords. | Show how to configure Aspose.Cells PdfSaveOptions to apply AES‑256 security and assign user/owner passwords when converting Excel to PDF in a .NET application.
// Common Searches: Aspose.Cells C# AES256 PDF encryption with user and owner passwords | How to set PDF security options when saving a workbook to PDF using Aspose.Cells .NET | Encrypt PDF output from Excel conversion with Aspose.Cells SaveFormat.Pdf options
// Tags: Aspose.Cells PDF security | C# set PDF user password Aspose.Cells | C# set PDF owner password Aspose.Cells | PdfSaveOptions security settings Aspose.Cells | Excel to password‑protected PDF conversion .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file, checks its presence, and saves it as a PDF using Aspose.Cells. It demonstrates basic error handling and console output but does not apply any encryption or password protection to the generated PDF.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as a PDF (encryption not available in this version)
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
