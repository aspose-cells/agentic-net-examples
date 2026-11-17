// Title: Generate a PDF with a separate page for each Excel worksheet using Aspose.Cells PdfSaveOptions in C#
// Common Searches: C# Aspose.Cells export Excel workbook to PDF with one page per sheet | How to enable OnePagePerSheet in PdfSaveOptions for PDF conversion with Aspose.Cells | Saving an Excel file as PDF where each worksheet becomes its own page using Aspose.Cells | Aspose.Cells PdfSaveOptions configuration for per‑worksheet PDF pagination in .NET | C# code example for PdfSaveOptions.OnePagePerSheet property
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | C# Excel to PDF per worksheet | PDF pagination per Excel sheet Aspose | .NET export workbook to PDF with separate pages | Aspose.Cells PDF conversion settings

using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, sets PdfSaveOptions.OnePagePerSheet to true, and saves it as a PDF so that each worksheet appears on its own PDF page.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to generate a separate PDF page for each worksheet
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = true;

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
