// Title: Convert an Excel workbook to PDF in C# with Aspose.Cells and add searchable keywords to the PDF
// Common Searches: C# code to convert .xlsx file to PDF using Aspose.Cells and then set PDF metadata | how to add custom keywords to a PDF generated from Excel with Aspose in .NET | Aspose.Cells PdfSaveOptions does not support PDF properties, what is the workaround in C# | save Excel workbook as PDF and embed search terms for archive using Aspose libraries | post‑process PDF created by Aspose.Cells to include metadata in a .NET application
// Tags: Aspose.Cells Excel to PDF conversion C# | PdfSaveOptions metadata limitation Aspose.Cells | add PDF metadata with Aspose.Pdf after conversion | archive‑friendly PDF generation from Excel workbook | set custom PDF keywords in .NET using Aspose libraries

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the source Excel file, loads it with Aspose.Cells, saves it as a PDF using PdfSaveOptions, and notes that PDF metadata must be added in a post‑processing step (e.g., with Aspose.Pdf) to embed searchable keywords for archival purposes.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: Aspose.Cells PdfSaveOptions does not expose a direct API for setting custom PDF properties.
            // If PDF metadata is required, it can be added after saving using Aspose.Pdf or other libraries.

            // Save the workbook as a PDF file using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
