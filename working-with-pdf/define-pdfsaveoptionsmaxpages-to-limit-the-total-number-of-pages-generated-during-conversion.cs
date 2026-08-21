// Title: Limit PDF Pages with PdfSaveOptions.PageCount in AspNet Aspose.Cells
// Description: Learn how to restrict the number of pages generated when converting an Excel workbook to PDF using Aspose.Cells for .NET. The example shows setting PdfSaveOptions.PageCount to export only the first N pages.
// Keywords: Aspose.Cells PDF page limit | PdfSaveOptions PageCount C# | Excel to PDF max pages | Aspose.Cells .NET example | GitHub Aspose.Cells PDF conversion
// Common Searches: Aspose.Cells set maximum PDF pages | PdfSaveOptions.PageCount usage | C# limit pages when saving Excel as PDF | How to export only first few pages with Aspose.Cells | Aspose.Cells PDF pagination control
// Developer Intent: Configure a ceiling on the page count of a PDF produced from an Excel file.
// Use Cases: Generate a quick preview PDF containing only the first few pages for stakeholder review. | Create a summary document from a large workbook while keeping file size low. | Enforce page‑count limits in a web service that converts user‑uploaded Excel files to PDF.
// AI Prompts: Show a C# snippet that uses Aspose.Cells PdfSaveOptions.PageCount to export the first N pages of a workbook. | Explain how PageCount interacts with worksheet PageSetup settings and how to determine the appropriate value. | Provide code that calculates PageCount dynamically based on row count, column count, or user input.

using System;
using Aspose.Cells;

// Learn how to restrict the number of pages generated when converting an Excel workbook to PDF using Aspose.Cells for .NET. The example shows setting PdfSaveOptions.PageCount to export only the first N pages.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Limit the total number of pages generated during conversion
        // (PageCount acts as the maximum pages to save)
        pdfOptions.PageCount = 5; // adjust the value as needed

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);

        Console.WriteLine("PDF saved with a maximum of 5 pages.");
    }
}
