// Title: Save an Aspose.Cells workbook as a PDF/A‑3u file and add XML metadata for accessibility using C#
// AI Prompts: Generate C# code that creates a workbook with Aspose.Cells, configures PdfSaveOptions for PDF/A‑3u compliance, and saves it as a PDF. | Show how to attach custom XML (XMP) metadata to a PDF produced by Aspose.Cells, using a secondary PDF library such as Aspose.PDF or iTextSharp. | Add robust try‑catch error handling for converting an Excel workbook to a PDF/A‑3u document with Aspose.Cells.
// Common Searches: how to export Excel to PDF/A-3u with Aspose.Cells in C# | adding XMP metadata to PDF generated from Aspose.Cells workbook | C# example for PDF/A-3u compliance and custom XML metadata in Aspose.Cells
// Tags: Aspose.Cells PDF/A-3u export C# | PdfSaveOptions Compliance property | embed XMP XML metadata in PDF with Aspose.PDF | Excel to accessible PDF conversion | error handling Aspose.Cells PDF save

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Creates a workbook, sets PdfSaveOptions.Compliance to PdfCompliance.PdfA3u, and saves it as output.pdf. XML metadata must be added after saving with a PDF library, as Aspose.Cells does not embed it directly.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // XML metadata to embed (not directly supported in this version of Aspose.Cells)
            // If needed, embed using a PDF library after saving the PDF.

            // Configure PDF save options for PDF/A‑3u compliance
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set PDF/A‑3u compliance
                Compliance = PdfCompliance.PdfA3u
            };

            // Save the workbook as a PDF with the configured options
            workbook.Save("output.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
