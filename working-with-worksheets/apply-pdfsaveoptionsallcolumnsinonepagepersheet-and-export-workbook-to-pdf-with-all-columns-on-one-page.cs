// Title: How to export an Excel workbook to PDF with all columns fitting on one page per sheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to PDF with all columns on one page per sheet | PdfSaveOptions AllColumnsInOnePagePerSheet usage example in .NET | Fit all worksheet columns onto a single PDF page using Aspose.Cells | C# code to save workbook as PDF with columns scaled to one page | Configure Aspose.Cells PDF export to fit columns per sheet
// Tags: Aspose.Cells PdfSaveOptions column fit | C# Excel to PDF conversion with column scaling | AllColumnsInOnePagePerSheet configuration | single-page PDF export per worksheet | fit worksheet columns to PDF using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example loads an Excel workbook, sets the PdfSaveOptions.AllColumnsInOnePagePerSheet flag to true, and saves the file as a PDF so that each worksheet's columns are compressed to fit on a single page.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to fit all columns on one page per sheet
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AllColumnsInOnePagePerSheet = true
        };

        // Export the workbook to PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
