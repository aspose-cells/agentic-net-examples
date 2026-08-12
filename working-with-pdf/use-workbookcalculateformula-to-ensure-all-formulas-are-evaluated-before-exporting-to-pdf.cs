// Title: C# – Use Workbook.CalculateFormula to evaluate formulas before Aspose.Cells PDF export
// Description: Shows how to force Aspose.Cells to compute every worksheet formula with Workbook.CalculateFormula, then save the workbook as a PDF using PdfSaveOptions (CalculateFormula = true). The resulting PDF contains the evaluated values.
// Keywords: Aspose.Cells C# calculate formulas | Workbook.CalculateFormula example | PdfSaveOptions CalculateFormula property | export Excel to PDF C# | evaluate formulas before PDF export | Aspose.Cells PDF generation | C# spreadsheet to PDF
// Common Searches: Aspose.Cells calculate all formulas before PDF export | Workbook.CalculateFormula vs PdfSaveOptions.CalculateFormula C# | How to ensure formulas are evaluated when saving Excel as PDF with Aspose.Cells | C# export Excel workbook to PDF with evaluated formulas | Aspose.Cells PDF export formula evaluation
// Developer Intent: The developer needs to guarantee that every formula in a workbook is calculated before the file is saved as a PDF.
// Use Cases: Financial statements PDF where totals must reflect the latest calculations. | Invoice PDFs that include tax, discount, and subtotal formulas. | Automated reporting pipelines that convert Excel sheets to PDF with up‑to‑date derived values.
// AI Prompts: Write C# code using Aspose.Cells to calculate all workbook formulas and then export the workbook to PDF with the results embedded. | Explain the impact of setting PdfSaveOptions.CalculateFormula to true when Workbook.CalculateFormula has already been called. | Provide a C# example with error handling for exporting a formula‑rich workbook to PDF using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to force Aspose.Cells to compute every worksheet formula with Workbook.CalculateFormula, then save the workbook as a PDF using PdfSaveOptions (CalculateFormula = true). The resulting PDF contains the evaluated values.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data and a formula
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

        // Explicitly calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Set PDF save options (optional: CalculateFormula can be true, but formulas are already evaluated)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CalculateFormula = true
        };

        // Export the workbook to PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
