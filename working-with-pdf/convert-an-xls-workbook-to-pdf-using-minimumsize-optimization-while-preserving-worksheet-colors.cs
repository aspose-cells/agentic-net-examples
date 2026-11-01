// Title: Convert an XLS workbook to PDF with MinimumSize optimization while keeping worksheet colors using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells convert XLS to PDF with minimum file size | How to retain Excel worksheet colors when saving as PDF using Aspose.Cells | PdfSaveOptions MinimumSize example for Excel to PDF conversion in .NET | Reduce PDF size from XLS workbook with Aspose.Cells while preserving formatting | Aspose.Cells PDF export preserving colors and optimizing for smallest size
// Tags: Aspose.Cells PDFSaveOptions MinimumSize | XLS to PDF conversion using Aspose.Cells C# | retain worksheet colors PDF export | minimum size PDF optimization Aspose.Cells | Excel to PDF color preservation .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The code loads an XLS workbook, configures PdfSaveOptions with OptimizationType set to MinimumSize to shrink the output, and saves the workbook as a PDF while automatically preserving the original worksheet colors.
class Program
{
    static void Main()
    {
        // Load the source XLS workbook
        // (Replace "input.xls" with the actual path to your XLS file)
        Workbook workbook = new Workbook("input.xls");

        // Configure PDF save options for MinimumSize optimization
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Optimize the PDF for minimum file size
            OptimizationType = PdfOptimizationType.MinimumSize,

            // Preserve the original worksheet colors (default behavior)
            // No additional settings required; colors are retained automatically
        };

        // Save the workbook as PDF using the configured options
        // (Replace "output.pdf" with the desired output file path)
        workbook.Save("output.pdf", pdfOptions);
    }
}
