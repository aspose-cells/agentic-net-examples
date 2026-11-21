// Title: Convert an XLS workbook to PDF with MinimumSize optimization while preserving worksheet colors using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# convert XLS to PDF with MinimumSize optimization | preserve Excel worksheet colors when exporting to PDF using Aspose.Cells | reduce PDF size from Excel workbook using PdfSaveOptions MinimumSize in .NET | C# example saving XLS as PDF while keeping cell colors and minimizing file size | Aspose.Cells PDFSaveOptions settings to retain colors and achieve smallest PDF
// Tags: PdfSaveOptions MinimumSize Aspose.Cells | XLS to PDF conversion with color preservation | Excel worksheet color retention in PDF export | PDF file size optimization using Aspose.Cells | C# Aspose.Cells PDF export with size reduction

using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an XLS workbook, applies PdfSaveOptions with MinimumSize optimization, and saves it as a PDF while automatically retaining the original worksheet colors.
class Program
{
    static void Main()
    {
        // Load the source XLS workbook
        Workbook workbook = new Workbook("input.xls");

        // Configure PDF save options for MinimumSize optimization
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use the MinimumSize optimization to reduce file size
            OptimizationType = PdfOptimizationType.MinimumSize,
            // Preserve worksheet colors (default behavior, but explicitly keep them)
            // No additional settings needed as colors are retained automatically
        };

        // Save the workbook as a PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
