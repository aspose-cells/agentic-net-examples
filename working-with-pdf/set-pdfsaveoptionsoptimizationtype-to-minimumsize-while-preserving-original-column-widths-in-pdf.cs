// Title: Export an Excel workbook to PDF with MinimumSize optimization while keeping original column widths using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set PdfSaveOptions OptimizationType to MinimumSize | How to keep column widths when saving Excel as PDF with Aspose.Cells | Reduce PDF file size from Excel conversion using Aspose.Cells PdfSaveOptions | C# export XLSX to PDF with smallest file size Aspose.Cells | PdfSaveOptions MinimumSize effect on column layout Aspose.Cells
// Tags: PdfSaveOptions MinimumSize optimization | Excel to PDF column width preservation | Aspose.Cells PDF export size reduction | C# Aspose.Cells PDF conversion settings | Workbook.Save PDF with optimization

using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, configures PdfSaveOptions with OptimizationType set to MinimumSize to minimize the PDF size, relies on the default behavior that retains original column widths, and saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set the optimization type to MinimumSize to reduce file size
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Column widths are preserved by default when saving to PDF,
        // so no additional settings are required.

        // Save the workbook as a PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
