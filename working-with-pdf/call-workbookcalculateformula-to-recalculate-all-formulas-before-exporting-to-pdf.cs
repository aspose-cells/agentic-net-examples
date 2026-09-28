// Title: Recalculate all formulas in an Excel workbook with Aspose.Cells for .NET before saving as PDF
// AI Prompts: Load an .xlsx file with Aspose.Cells, recalculate all embedded formulas, and generate a PDF output in C#. | Illustrate invoking a full formula refresh before calling the Save method to produce a PDF with Aspose.Cells. | Write a C# example that ensures Excel calculations are up‑to‑date and then converts the workbook to PDF using Aspose.Cells.
// Common Searches: Aspose.Cells .NET recalculate formulas before PDF export | C# force Excel formula evaluation with Aspose.Cells prior to saving as PDF | Workbook.CalculateFormula example for PDF generation in Aspose.Cells
// Tags: Workbook.CalculateFormula usage | Excel to PDF conversion after formula refresh | Aspose.Cells formula recalculation before export | C# PDF generation with updated Excel calculations | force formula evaluation Aspose.Cells .NET

using Aspose.Cells;

// // Loads an Excel workbook, forces a full formula recalculation, and saves the workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
