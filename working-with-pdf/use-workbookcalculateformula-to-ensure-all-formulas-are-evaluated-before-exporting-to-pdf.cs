// Title: How to recalculate all formulas in an Excel workbook before exporting to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Load an .xlsx file, trigger a full formula evaluation using Workbook.CalculateFormula, and then generate a PDF with Aspose.Cells in C#. | Write C# code that refreshes every Excel calculation before calling Workbook.Save with SaveFormat.Pdf. | Show how to force a complete formula refresh and export the workbook to PDF using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# recalculate workbook formulas prior to PDF conversion | Force formula update when saving Excel as PDF with Aspose.Cells | C# example for Workbook.CalculateFormula followed by PDF export | Ensure PDF output reflects latest Excel calculations using Aspose.Cells .NET
// Tags: formula calculation API Aspose.Cells | refresh workbook calculations Aspose.Cells | Excel to PDF output with updated formulas | ensure formulas are evaluated prior to PDF generation | Aspose.Cells C# PDF generation after calculations

using Aspose.Cells;

// Loads an Excel workbook, forces full formula calculation with Workbook.CalculateFormula, and saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure all formulas in the workbook are evaluated
        workbook.CalculateFormula();

        // Export the workbook to PDF after formulas have been calculated
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
