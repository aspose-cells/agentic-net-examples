// Title: Limit PDF output to the first 10 pages when saving an Aspose.Cells workbook in C#
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to PDF using PdfSaveOptions with a maximum of 10 pages. | Show how to set PageIndex and PageCount in PdfSaveOptions to export only the first ten pages of a workbook. | Demonstrate configuring Aspose.Cells PDF export to restrict the generated PDF to ten pages.
// Common Searches: how to export only first 10 pages of Excel to PDF using Aspose.Cells C# | Aspose.Cells PdfSaveOptions limit number of pages .NET | C# specify maximum pages when saving workbook to PDF with Aspose.Cells | how to cap PDF output at ten pages using Aspose.Cells | Aspose.Cells PDF export page range example C#
// Tags: Aspose.Cells PDF export page limit | PdfSaveOptions page limit C# | limit PDF pages Aspose.Cells | export workbook to PDF first ten pages | C# Aspose.Cells PDF page range

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program creates a workbook, populates it with sample data, and saves it as a PDF limited to the first ten pages by configuring PdfSaveOptions with PageIndex = 0 and PageCount = 10.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate the worksheet with sample data to generate multiple pages
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 200; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Configure PDF save options to limit the output to the first 10 pages
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Zero‑based index of the first page to render
                PageIndex = 0,
                // Number of pages to render starting from PageIndex
                PageCount = 10
            };

            // Save the workbook as PDF using the configured options
            workbook.Save("LimitedOutput.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
