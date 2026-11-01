// Title: Convert a password‑protected XLS workbook to PDF with embedded fonts using Aspose.Cells for .NET
// Common Searches: c# load password protected xls with Aspose.Cells and save as pdf | asp.net convert protected excel file to pdf with font embedding | how to use LoadOptions password property in Aspose.Cells C# | export excel workbook to pdf with embedded fonts for printing using Aspose.Cells | read encrypted xls file in C# and generate printable pdf | Aspose.Cells PdfSaveOptions default font embedding behavior
// Tags: Aspose.Cells password protected XLS loading | PDF export with embedded fonts Aspose.Cells | C# Excel to PDF conversion Aspose.Cells | LoadOptions password property usage | PdfSaveOptions default font embedding

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the presence of a password‑protected XLS file, loads it into an Aspose.Cells Workbook using LoadOptions with the specified password, and then saves the workbook as a PDF. The default PdfSaveOptions embed the necessary fonts, producing a printable PDF while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        // Input XLS file (password‑protected)
        string inputFile = "protected.xls";

        // Password for the workbook
        string password = "myPassword";

        // Output PDF file
        string outputFile = "output.pdf";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputFile))
                throw new FileNotFoundException($"Input file not found: {inputFile}");

            // Load the workbook with the password (auto‑detect format)
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Configure PDF save options (default settings are sufficient)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputFile, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
