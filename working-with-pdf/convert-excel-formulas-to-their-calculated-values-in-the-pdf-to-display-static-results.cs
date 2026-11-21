// Title: Export Excel workbook to PDF with evaluated formulas as static values using Aspose.Cells for .NET
// Common Searches: how to export an Excel file to PDF with calculated results using Aspose.Cells C# | Aspose.Cells C# render formulas as values when saving to PDF | force formula evaluation before PDF conversion in .NET with Aspose.Cells | save Excel worksheet as PDF showing static values not formulas Aspose | configure PdfSaveOptions to calculate formulas during PDF export in C#
// Tags: calculate formulas before PDF export Aspose.Cells | static values in PDF using Aspose.Cells | PdfSaveOptions CalculateFormula property C# | one page per worksheet PDF Aspose.Cells | export Excel to PDF with evaluated formulas

using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, forces all formulas to be calculated, sets PdfSaveOptions to evaluate formulas and optionally place each worksheet on a separate PDF page, then saves the workbook as a PDF containing static calculated results.
class Program
{
    static void Main()
    {
        // Load the Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Force calculation of all formulas so that cells contain static values
        workbook.CalculateFormula();

        // Configure PDF save options to render calculated values
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure formulas are evaluated before rendering
            CalculateFormula = true,
            // Optional: one PDF page per worksheet
            OnePagePerSheet = true
        };

        // Save the workbook as a PDF with static results
        workbook.Save("output.pdf", pdfOptions);
    }
}
