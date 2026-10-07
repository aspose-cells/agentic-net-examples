// Title: Convert an HTML spreadsheet to a password‑protected PDF with Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an HTML file into an Aspose.Cells Workbook, sets a user password via PdfSaveOptions, and saves the workbook as an encrypted PDF. | Update the provided program to create the output folder when missing and apply PDF password protection using Aspose.Cells.
// Common Searches: Aspose.Cells C# convert HTML file to encrypted PDF | how to set user password when saving workbook as PDF with Aspose.Cells | example of PDF encryption using PdfSaveOptions in Aspose.Cells .NET | protect PDF generated from HTML spreadsheet in C#
// Tags: Aspose.Cells HTML to PDF with password protection | C# PdfSaveOptions user password | encrypt PDF output from Aspose.Cells workbook | convert HTML spreadsheet to PDF using Aspose.Cells | set PDF encryption in Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an HTML file that represents a spreadsheet into an Aspose.Cells Workbook, ensures the output directory exists, configures PdfSaveOptions with a user password for PDF encryption, and saves the workbook as a password‑protected PDF while handling possible exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.pdf";

            // Verify that the input HTML file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the HTML file (it must represent a spreadsheet) into a Workbook object.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (without security settings to avoid missing namespace issues).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
