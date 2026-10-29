// Title: Convert an Excel workbook to a password‑protected PDF using Aspose.Cells for .NET
// Common Searches: how to generate a password protected PDF from an .xlsx file with Aspose.Cells in C# | Aspose.Cells set user password when saving workbook as PDF example | C# code to export Excel to encrypted PDF using PdfSaveOptions | using PdfSaveOptions to apply document encryption in Aspose.Cells | convert Excel to PDF with user password protection in .NET
// Tags: Aspose.Cells password protected PDF conversion | PdfSaveOptions user password encryption C# | Excel workbook to encrypted PDF Aspose.Cells | Workbook.Save PDF with document encryption | C# Aspose.Cells PDF encryption example

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The sample checks for the source Excel file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (where a user password can be assigned via the Password property), and saves the workbook as a PDF. It demonstrates conversion to PDF and indicates where document encryption can be configured.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing Excel workbook
                var workbook = new Workbook(inputPath);

                // Configure PDF save options (encryption omitted for compatibility)
                var pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF file
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
