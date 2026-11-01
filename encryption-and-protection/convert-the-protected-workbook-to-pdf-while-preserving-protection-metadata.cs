// Title: Convert a password‑protected Excel workbook to PDF with Aspose.Cells for .NET while retaining protection metadata
// Common Searches: how to load a password protected XLSX and save as PDF using Aspose.Cells C# | Aspose.Cells export encrypted Excel workbook to PDF preserving protection | C# code to convert protected Excel file to PDF with Aspose.Cells | save password‑protected workbook to PDF without losing protection info Aspose.Cells | Aspose.Cells .NET load workbook with password and generate PDF
// Tags: Aspose.Cells load password protected workbook | convert protected Excel to PDF C# | PdfSaveOptions usage Aspose.Cells | preserve workbook protection metadata Aspose.Cells | encrypted XLSX to PDF conversion .NET

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates loading a password‑protected XLSX workbook using LoadOptions with the password, then saving it as a PDF via PdfSaveOptions. Note that Aspose.Cells currently does not expose a setting to retain workbook‑protection metadata in the generated PDF.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the protected Excel workbook
            string inputPath = "protected.xlsx";

            // Password used to protect the workbook
            string password = "myPassword";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook with the specified password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Configure PDF save options (no PreserveWorkbookProtection property in current API)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            string outputPath = "protected_output.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
