// Title: How to export an Excel workbook to PDF with solid gridlines using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates through all worksheets, enables gridline printing, and saves the workbook as a PDF with Aspose.Cells. | Show how to configure PdfSaveOptions so that gridlines are preserved during XLSX‑to‑PDF conversion in a .NET application. | Provide a sample that sets each worksheet’s gridline visibility and generates a PDF containing solid lines using Aspose.Cells.
// Common Searches: Aspose.Cells C# export Excel to PDF with visible gridlines | Enable solid gridlines when converting XLSX to PDF using Aspose.Cells .NET | PrintGridlines property effect on PDF output in Aspose.Cells | How to keep gridlines in PDF generated from workbook with Aspose.Cells | PdfSaveOptions settings for preserving Excel gridlines in PDF
// Tags: enable worksheet gridline printing Aspose.Cells | PDF conversion retaining Excel gridlines | PdfSaveOptions preserve gridlines Aspose.Cells | C# solid gridlines PDF conversion Aspose.Cells | iterate worksheets set gridline printing

using Aspose.Cells;
using System;

// Loads an XLSX workbook, enables gridline printing for every worksheet, optionally adjusts PdfSaveOptions, and saves the workbook as a PDF where the gridlines appear as solid lines.
class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Enable printing of gridlines for every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.PageSetup.PrintGridlines = true;
        }

        // Configure PDF save options (optional adjustments)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: keep each worksheet on its own page
            OnePagePerSheet = false
        };

        // Save the workbook as a PDF with solid gridlines rendered
        workbook.Save("output.pdf", pdfOptions);
    }
}
