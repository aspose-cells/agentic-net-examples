// Title: Recalculate all Excel formulas with Aspose.Cells before saving the workbook as PDF in C#
// Common Searches: Aspose.Cells C# recalculate all workbook formulas before PDF export | How to force formula calculation when saving Excel to PDF using Aspose.Cells | C# code to update Excel formulas then export to PDF with Aspose.Cells | Workbook.CalculateFormula followed by SaveFormat.Pdf example in .NET | Programmatically refresh Excel calculations prior to PDF conversion with Aspose.Cells
// Tags: Workbook.CalculateFormula C# | export workbook to PDF after formula recalculation | Aspose.Cells PDF conversion with updated formulas | recalculate formulas before SaveFormat.Pdf | force formula evaluation in .NET Excel to PDF

using Aspose.Cells;

// The example loads an Excel file into an Aspose.Cells Workbook, calls Workbook.CalculateFormula to recompute every formula, and then saves the workbook as a PDF using SaveFormat.Pdf, ensuring the exported PDF reflects the latest calculations.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Export the workbook to PDF after recalculation
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
