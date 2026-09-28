// Title: Generate a PDF with evaluated Excel formulas as static values using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, forces all formulas to be evaluated, and saves the workbook as a PDF with the calculated results displayed as static text using Aspose.Cells. | Show how to configure Aspose.Cells PdfSaveOptions to fit all columns on a single page per sheet while preserving the evaluated formula values in the exported PDF. | Demonstrate the correct order of calling Workbook.CalculateFormula followed by Workbook.Save with PdfSaveOptions to ensure formulas are converted to values in the resulting PDF.
// Common Searches: asp.net how to export Excel to PDF with formulas evaluated using Aspose.Cells | c# convert Excel formulas to static values before saving as PDF | Aspose.Cells PdfSaveOptions fit columns on one page per sheet example | force calculation of all formulas in workbook prior to PDF conversion Aspose.Cells
// Tags: calculate formulas Aspose.Cells Workbook | export workbook to PDF Aspose.Cells | PdfSaveOptions AllColumnsInOnePagePerSheet | static values in PDF Aspose.Cells | C# evaluate Excel formulas before PDF export

using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads an Excel workbook, forces calculation of every formula, sets PDF options to fit all columns on one page per sheet, and saves the workbook as a PDF that displays the evaluated results as static values.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Force calculation of all formulas so that static values are stored
        workbook.CalculateFormula();

        // Configure PDF save options (default behavior already writes calculated values)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Example option: fit all columns on one page per sheet (optional)
        pdfOptions.AllColumnsInOnePagePerSheet = true;

        // Save the workbook as a PDF; the PDF will display the calculated results
        workbook.Save("output.pdf", pdfOptions);
    }
}
