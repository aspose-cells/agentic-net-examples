// Title: How to export an Excel workbook to PDF with one page per sheet and fit all columns using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to PDF with OnePagePerSheet true | Fit all columns on a single PDF page per worksheet using Aspose.Cells | PdfSaveOptions AllColumnsInOnePagePerSheet example in C# | Generate single-page PDF per sheet from .xlsx with Aspose.Cells .NET
// Tags: PdfSaveOptions single page per sheet | fit all columns on one PDF page | Aspose.Cells Excel to PDF export settings | C# PDF export with column scaling | Aspose.Cells .NET PDF generation options

using Aspose.Cells;
using System;

// The example loads an Excel workbook, configures PdfSaveOptions with OnePagePerSheet and AllColumnsInOnePagePerSheet set to true, and saves the workbook as a PDF where each worksheet is rendered on a single page with all columns compressed to fit that page.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure each worksheet is rendered on a single PDF page
            OnePagePerSheet = true,
            // Fit all columns of the worksheet onto that single page
            AllColumnsInOnePagePerSheet = true
        };

        // Save the workbook as a PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
