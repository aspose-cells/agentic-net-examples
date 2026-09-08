// Title: Convert an XLS workbook to PDF with MinimumSize optimization while preserving worksheet colors using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xls file and saves it as a PDF using Aspose.Cells with PdfSaveOptions configured for MinimumSize to achieve the smallest possible file. | Show how to keep original worksheet background colors intact when exporting an Excel workbook to PDF with Aspose.Cells in a .NET application. | Explain the steps to set PdfOptimizationType.MinimumSize in Aspose.Cells and generate a PDF without altering the workbook's formatting.
// Common Searches: asp.net convert xls to pdf with minimum file size using aspose.cells | c# preserve cell colors when exporting excel to pdf with aspose | pdfsaveoptions minimumsize optimization asp.net example | how to reduce pdf size from excel conversion asp.net aspose.cells
// Tags: Aspose.Cells PDF minimum size optimization | XLS to PDF conversion preserving colors | PdfSaveOptions OptimizationType MinimumSize C# | Excel worksheet color retention in PDF export | Aspose.Cells PDF export file size reduction

using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads an XLS workbook, applies PdfSaveOptions with OptimizationType.MinimumSize to minimize PDF size while retaining worksheet colors, and saves the result as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the source XLS workbook
        Workbook workbook = new Workbook("input.xls");

        // Configure PDF save options for MinimumSize optimization
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use the smallest file size optimization
            OptimizationType = PdfOptimizationType.MinimumSize,
            // Preserve worksheet colors (default behavior, no extra setting required)
            // Additional settings can be added here if needed
        };

        // Save the workbook as a PDF file with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
