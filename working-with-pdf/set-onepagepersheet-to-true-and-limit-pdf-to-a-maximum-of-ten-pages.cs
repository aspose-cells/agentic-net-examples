// Title: Set OnePagePerSheet to true and limit PDF output to ten pages with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set OnePagePerSheet true when saving workbook to PDF | How to restrict PDF output to ten pages using Aspose.Cells .NET | Export Excel workbook to PDF with one page per worksheet and page count limit in C# | Aspose.Cells PDF save options to control maximum number of pages | Trim or remove extra pages after converting Excel to PDF with Aspose.Cells
// Tags: OnePagePerSheet PDF save option Aspose.Cells | PDF page limit handling Aspose.Cells .NET | single page per worksheet export Aspose.Cells | post‑generation PDF trimming Aspose.Cells | Aspose.Cells PDF pagination control

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, fills it with sample data, configures PdfSaveOptions to render each worksheet on a single page by setting OnePagePerSheet to true, notes that Aspose.Cells does not provide a direct MaxPages property, and saves the workbook as a PDF. To enforce a ten‑page maximum, additional post‑processing of the generated PDF is required.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";
            for (int i = 0; i < 200; i++)
            {
                sheet.Cells[i, 0].PutValue($"Row {i}");
                sheet.Cells[i, 1].PutValue(i);
            }

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Render each worksheet on a single page
                OnePagePerSheet = true
                // Note: Aspose.Cells PdfSaveOptions does not provide a MaxPages property.
                // If page limiting is required, it must be handled after PDF generation.
            };

            // Save the workbook as PDF
            workbook.Save("output.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
