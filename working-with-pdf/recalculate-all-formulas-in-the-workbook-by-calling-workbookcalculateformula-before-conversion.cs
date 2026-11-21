// Title: Recalculate all Excel formulas with Workbook.CalculateFormula and export the workbook to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells calculate all formulas before saving workbook as PDF | How to force formula evaluation in Aspose.Cells prior to PDF export | Aspose.Cells .NET recalculate workbook formulas then convert to PDF | Workbook.CalculateFormula example for PDF conversion in C# | Export Excel to PDF with updated formula results using Aspose.Cells | Aspose.Cells calculate formulas on load before PDF generation | C# code to recalc Excel formulas and save as PDF with Aspose.Cells
// Tags: calculate formulas Aspose.Cells C# | PDF export after formula evaluation Aspose.Cells | Aspose.Cells recalculate workbook before save | Excel to PDF conversion with updated formulas C# | Aspose.Cells formula refresh prior to PDF generation | C# workbook calculation and PDF conversion

using Aspose.Cells;

// The example loads an Excel file, forces a full formula recalculation using Workbook.CalculateFormula, and then saves the workbook as a PDF document with the updated values.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook after calculation (example: convert to PDF)
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
