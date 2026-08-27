// Title: Convert an Excel workbook to a grayscale PDF with minimal file size using Aspose.Cells for .NET
// AI Prompts: Write C# code that saves a Workbook as a grayscale PDF while applying the smallest possible file size settings with Aspose.Cells. | Demonstrate how to set Worksheet.PageSetup.BlackAndWhite and use PdfSaveOptions.OptimizationType = PdfOptimizationType.MinimumSize for PDF export. | Provide a complete example that creates a workbook, adds data, and generates a compact black‑and‑white PDF file.
// Common Searches: asp.net convert excel to black and white PDF with Aspose.Cells | how to reduce PDF size when exporting Excel using Aspose.Cells C# | set grayscale rendering for PDF output in Aspose.Cells | PdfOptimizationType.MinimumSize example Aspose.Cells | export worksheet as PDF with black‑and‑white page setup C#
// Tags: grayscale PDF export Aspose.Cells | Worksheet.PageSetup.BlackAndWhite property | PdfOptimizationType.MinimumSize usage | Aspose.Cells PDF size optimization | export Excel to black‑and‑white PDF C#

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Creates a workbook, adds sample data, enables black‑and‑white page rendering, configures PdfSaveOptions for minimum size, and saves the file as GrayscaleOutput.pdf.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data (optional, just to have content)
        worksheet.Cells["A1"].PutValue("Sample Data for Grayscale PDF");

        // Enable black‑and‑white (grayscale) rendering for printing
        worksheet.PageSetup.BlackAndWhite = true;

        // Configure PDF save options to minimize file size
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF using the specified options
        workbook.Save("GrayscaleOutput.pdf", pdfOptions);
    }
}
