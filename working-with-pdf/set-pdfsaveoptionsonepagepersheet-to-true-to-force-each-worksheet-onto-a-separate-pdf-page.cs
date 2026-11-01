// Title: Force each worksheet onto a separate PDF page when converting Excel to PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as a PDF where each worksheet is placed on a distinct PDF page using Aspose.Cells. | Demonstrate how to configure PdfSaveOptions in Aspose.Cells to produce a PDF with a separate page for every worksheet. | Provide a complete C# example that sets the PDF export option to render each sheet on its own page and saves the workbook.
// Common Searches: Aspose.Cells C# export Excel workbook to PDF with each sheet on a new page | How to enable separate page per worksheet in PdfSaveOptions for PDF conversion .NET | Generate PDF where each Excel worksheet gets its own page using Aspose.Cells | C# example converting multi-sheet Excel to PDF with one page per sheet
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | C# Excel to PDF conversion per worksheet | force separate PDF page per sheet Aspose.Cells | configure PDF export options Aspose.Cells .NET | multi-sheet workbook PDF pagination Aspose

using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, sets PdfSaveOptions.OnePagePerSheet to true so each worksheet is rendered on its own PDF page, and saves the result as a PDF.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options and set OnePagePerSheet to true
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = true; // Force each worksheet onto a separate PDF page

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
