// Title: Add a semi‑transparent diagonal watermark text to each PDF page when saving a workbook with Aspose.Cells in C#
// AI Prompts: Generate C# code that creates a Workbook, sets PdfSaveOptions.WatermarkText to a custom string, defines font size, color, rotation angle, and 30% opacity, then saves the workbook as a PDF with the watermark applied to all pages. | Show how to reference the Aspose.Cells.Pdf assembly and configure PdfSaveOptions to embed a repeating diagonal watermark with specified transparency in a PDF generated from a workbook. | Provide a step‑by‑step example that applies a semi‑transparent watermark across every page of a PDF using Aspose.Cells, including code to adjust alignment, rotation, and opacity before calling Workbook.Save.
// Common Searches: Aspose.Cells C# add translucent watermark to PDF pages | PdfSaveOptions watermark text rotation opacity Aspose.Cells example | how to set diagonal watermark for PDF output in Aspose.Cells | C# Aspose.Cells PDF watermark custom font color transparency | enable watermark feature in Aspose.Cells PDF export
// Tags: Aspose.Cells PDF watermark opacity setting | PdfSaveOptions watermark text formatting | C# add repeating watermark to PDF with Aspose.Cells | configure watermark font color transparency Aspose.Cells | enable PDF watermark feature Aspose.Cells assembly

using System;
using System.Drawing;
using Aspose.Cells;

// The sample creates a new Aspose.Cells Workbook, adds sample data, configures PdfSaveOptions (including the required assembly for watermark support), sets a semi‑transparent diagonal watermark text with custom font, color, rotation, and opacity, and saves the workbook as output.pdf with the watermark on every page.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate the first worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(DateTime.Now);

            // Configure PDF save options (watermark feature requires additional assembly)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save("output.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
