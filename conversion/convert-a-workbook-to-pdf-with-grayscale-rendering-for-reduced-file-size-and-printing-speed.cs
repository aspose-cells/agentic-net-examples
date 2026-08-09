// Title: C# – Convert Excel Workbook to Grayscale PDF with Minimum File Size using Aspose.Cells
// Description: Create a workbook, enable PageSetup.BlackAndWhite, set PdfSaveOptions to MinimumSize and Flate compression, then save as a grayscale PDF for reduced size and faster printing.
// Keywords: Aspose.Cells | C# PDF conversion | grayscale PDF | BlackAndWhite PageSetup | PdfSaveOptions | MinimumSize optimization | Flate compression | Excel to PDF .NET | reduce PDF size | print speed
// Common Searches: Aspose.Cells export grayscale PDF | C# save Excel as black‑and‑white PDF | minimum size PDF Aspose.Cells | how to use PdfSaveOptions OptimizationType MinimumSize | apply Flate compression with Aspose.Cells | reduce PDF file size from Excel .NET
// Developer Intent: Produce a PDF from an Excel workbook in black‑and‑white mode while minimizing file size for faster printing or storage.
// Use Cases: Generating lightweight grayscale PDFs for bulk report printing | Archiving Excel data as compact PDFs to save storage | Creating compliance documents where color is prohibited | Automating batch conversion of workbooks to small PDFs in server‑side .NET applications
// AI Prompts: Write C# code that uses Aspose.Cells to convert a workbook to a grayscale PDF with MinimumSize optimization and Flate compression. | Explain the effect of PageSetup.BlackAndWhite on PDF output and how to combine it with PdfSaveOptions for smallest file size. | Provide a step‑by‑step tutorial for configuring Aspose.Cells PdfSaveOptions to produce a reduced‑size grayscale PDF in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Create a workbook, enable PageSetup.BlackAndWhite, set PdfSaveOptions to MinimumSize and Flate compression, then save as a grayscale PDF for reduced size and faster printing.
class ConvertWorkbookToPdfGrayscale
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Sample data for PDF conversion");

        // Enable black‑and‑white (grayscale) rendering for the worksheet
        // This forces all printed elements to be rendered in grayscale
        worksheet.PageSetup.BlackAndWhite = true;

        // Configure PDF save options for reduced file size
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Use the MinimumSize optimization type (smaller file, lower quality)
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Apply Flate compression to further shrink the PDF size
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;

        // Save the workbook as a PDF using the configured options
        workbook.Save("GrayscaleOutput.pdf", pdfOptions);
    }
}
