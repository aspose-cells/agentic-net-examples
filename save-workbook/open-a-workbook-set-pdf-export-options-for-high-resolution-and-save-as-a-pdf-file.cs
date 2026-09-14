// Title: Convert an Excel workbook to a high‑resolution PDF (300 DPI) with formula calculation using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, configures Aspose.Cells PdfSaveOptions to resample images at 300 DPI with 90 % JPEG quality, enables CalculateFormula, and saves the result as a PDF. | Show how to use Aspose.Cells in C# to produce a PDF that preserves formulas and renders images at high resolution from an Excel workbook.
// Common Searches: Aspose.Cells C# set PDF export DPI to 300 | Export Excel to PDF with high image quality using Aspose.Cells | C# convert .xlsx to PDF with formula calculation enabled | Adjust JPEG compression for PDF output in Aspose.Cells example
// Tags: image resample 300 DPI Aspose.Cells | high‑resolution PDF export Aspose.Cells | calculate formulas before PDF conversion C# | JPEG quality control for PDF Aspose.Cells | configure PDF save options Aspose.Cells

using System;
using Aspose.Cells;

// // Loads input.xlsx, sets PdfSaveOptions to resample images at 300 DPI with 90 % JPEG quality, enables formula calculation, and saves as output.pdf.
class Program
{
    static void Main()
    {
        // Load the existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set high resolution (e.g., 300 DPI) and JPEG quality (e.g., 90)
        pdfOptions.SetImageResample(300, 90);

        // Ensure formulas are calculated before saving
        pdfOptions.CalculateFormula = true;

        // Save the workbook as a PDF file using the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
