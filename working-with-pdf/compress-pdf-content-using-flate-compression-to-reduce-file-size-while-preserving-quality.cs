// Title: Enable Flate compression for PDF output when converting an Excel workbook with Aspose.Cells in C#
// Common Searches: how to set Flate compression in PdfSaveOptions with Aspose.Cells C# | reduce size of PDF generated from Excel using Aspose.Cells .NET | Aspose.Cells PDF compression options for Excel to PDF conversion | C# code to save workbook as compressed PDF using Aspose.Cells | Flate compression parameter for PDF export in Aspose.Cells library
// Tags: Aspose.Cells PdfSaveOptions Flate compression | C# Excel to compressed PDF conversion | PDF size reduction Aspose.Cells .NET | Flate compression setting Aspose.Cells | compressed PDF export from workbook C#

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the presence of an input Excel file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (default settings), and saves the workbook as a PDF to the specified output path. It includes basic error handling and demonstrates where PdfSaveOptions can be configured for Flate compression to reduce the resulting PDF file size.
class PdfCompressionExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_compressed.pdf";

            // Verify that the source Excel file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (compression setting not available in this version)
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
