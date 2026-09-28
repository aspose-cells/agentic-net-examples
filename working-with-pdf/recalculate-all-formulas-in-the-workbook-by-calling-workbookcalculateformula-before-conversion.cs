// Title: Recalculate all Excel formulas with Aspose.Cells and export the workbook to PDF in C#
// AI Prompts: Load an .xlsx file, force a full formula recalculation, and save the result as a PDF using Aspose.Cells for .NET. | Execute Workbook.CalculateFormula on a workbook before converting it to PDF with C#. | Refresh all calculated cells in an Excel workbook and generate a PDF output via Aspose.Cells.
// Common Searches: Aspose.Cells C# how to force formula evaluation before PDF conversion | Calculate all formulas in an Excel workbook then save as PDF using Aspose.Cells | C# workbook.CalculateFormula method usage for PDF export | Convert Excel to PDF after updating formulas with Aspose.Cells .NET
// Tags: Aspose.Cells formula refresh before PDF export | C# Excel workbook PDF generation after recalculation | recalculate cells then convert to PDF using Aspose | Excel to PDF workflow with updated calculations | trigger formula calculation in .NET workbook

using Aspose.Cells;

// Loads an Excel file, forces a full formula recalculation with Workbook.CalculateFormula, and saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        // Load the existing workbook (load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Recalculate all formulas in the workbook (required operation)
        workbook.CalculateFormula();

        // Save/convert the workbook after recalculation (save rule)
        // Example conversion to PDF; change format/path as needed
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
