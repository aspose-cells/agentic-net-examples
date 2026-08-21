// Title: Limit PDF Export to 10 Pages with Aspose.Cells PdfSaveOptions (C#)
// Description: Shows how to cap the PDF created from an Excel workbook at ten pages by setting the PdfSaveOptions.MaxPageCount property in Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | PdfSaveOptions | MaxPageCount | limit PDF pages | export Excel to PDF | PDF page limit | set maximum page count | Aspose.Cells PDF export | C# PDF generation
// Common Searches: Aspose.Cells limit PDF pages C# | PdfSaveOptions MaxPageCount example | how to restrict PDF page count with Aspose.Cells | C# export Excel to PDF with page limit | set maximum page count Aspose.Cells PDF
// Developer Intent: Configure Aspose.Cells to generate a PDF that contains no more than ten pages.
// Use Cases: Produce a short PDF preview of a large workbook for quick stakeholder review. | Create a concise report by capping the exported PDF at ten pages. | Generate a PDF sample for documentation while ensuring the file stays within a page limit.
// AI Prompts: Write C# code that loads an Excel file and saves it as a PDF limited to 10 pages using Aspose.Cells PdfSaveOptions.MaxPageCount. | Explain how to programmatically confirm that the saved PDF contains at most ten pages after using Aspose.Cells. | Suggest alternative techniques in Aspose.Cells to control PDF length, such as using PrintOptions or inserting manual page breaks.

using System;
using Aspose.Cells;

// Shows how to cap the PDF created from an Excel workbook at ten pages by setting the PdfSaveOptions.MaxPageCount property in Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Load an existing Excel workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Limit the PDF output to a maximum of 10 pages
        pdfOptions.PageCount = 10;

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
