// Title: Export each Excel worksheet to a single PDF page and fit all columns using Aspose.Cells PdfSaveOptions in C#
// AI Prompts: Write C# code that loads an .xlsx workbook, enables PdfSaveOptions.OnePagePerSheet and AllColumnsInOnePagePerSheet, and saves the workbook as a PDF with Aspose.Cells. | Show how to configure Aspose.Cells PDF conversion so that every worksheet is rendered on one page and all columns are automatically scaled to fit that page in a .NET application.
// Common Searches: Aspose.Cells C# export Excel to PDF with one page per worksheet and columns fit | PdfSaveOptions AllColumnsInOnePagePerSheet usage example | How to set OnePagePerSheet true for PDF output in Aspose.Cells | Fit all columns on a single PDF page when converting Excel with Aspose.Cells | C# code sample for single-page PDF per sheet using Aspose.Cells
// Tags: Aspose.Cells PDF export one page per sheet | PdfSaveOptions AllColumnsInOnePagePerSheet C# | Excel to PDF single worksheet page | scale columns to fit PDF page Aspose.Cells | configure PDF save options Aspose.Cells

using Aspose.Cells;
using System;

// Loads an Excel workbook, sets PdfSaveOptions.OnePagePerSheet and AllColumnsInOnePagePerSheet to true, and saves the workbook as a PDF where each worksheet occupies a single page with all columns scaled to fit.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Render each sheet on a single PDF page
        pdfOptions.OnePagePerSheet = true;

        // Fit all columns of the sheet onto that single page
        pdfOptions.AllColumnsInOnePagePerSheet = true;

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
