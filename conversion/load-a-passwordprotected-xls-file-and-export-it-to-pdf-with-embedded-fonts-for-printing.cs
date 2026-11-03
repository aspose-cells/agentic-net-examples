// Title: Convert a password‑protected XLS workbook to PDF with embedded fonts using Aspose.Cells for .NET
// Common Searches: c# load password protected xls file with Aspose.Cells LoadOptions | aspnet convert protected xls to pdf preserving fonts Aspose.Cells | how to embed fonts when saving workbook as pdf using Aspose.Cells | example code for opening encrypted Excel file and saving as PDF in C# | Aspose.Cells PDF export with default font embedding for printing
// Tags: Aspose.Cells LoadOptions password protection | Aspose.Cells PDF font embedding | XLS to PDF conversion Aspose.Cells .NET | encrypted Excel workbook loading Aspose.Cells | PDF export settings for printing Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the presence of a password‑protected XLS file, loads it into an Aspose.Cells Workbook using LoadOptions with the supplied password, configures PdfSaveOptions (which embed fonts by default), and saves the workbook as a PDF suitable for high‑quality printing, while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        // Path to the password‑protected XLS file
        string inputPath = "protected.xls";
        // Password for the file
        string password = "yourPassword";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook with the password (format is auto‑detected)
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Configure PDF save options (default options are sufficient for embedding fonts)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
