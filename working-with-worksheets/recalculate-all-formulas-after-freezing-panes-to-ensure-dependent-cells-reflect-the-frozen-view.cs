// Title: C# – Recalculate All Formulas After Freezing Panes with Aspose.Cells
// Description: Creates a workbook, adds values and formulas, freezes the top two rows and left two columns at cell C3, then calls Workbook.CalculateFormula() to refresh all dependent calculations before saving.
// Keywords: Aspose.Cells C# freeze panes | Workbook.CalculateFormula | recalculate formulas after FreezePanes | update dependent cells Excel .NET | freeze rows and columns Aspose.Cells | Excel formula refresh C# | Aspose.Cells worksheet calculation | freeze panes example .NET
// Common Searches: Aspose.Cells recalculate formulas after FreezePanes | C# freeze panes and calculate formulas | Workbook.CalculateFormula usage with FreezePanes | how to refresh formulas after freezing rows in Aspose.Cells | Aspose.Cells example freeze panes C#
// Developer Intent: Refresh every formula in a workbook so that cells dependent on frozen rows or columns show correct values after applying FreezePanes.
// Use Cases: Generate reports with frozen header rows where totals must reflect the latest data before export. | Maintain calculation consistency in large workbooks after programmatically freezing panes. | Prepare Excel files for end‑users that require both navigation aids (frozen panes) and accurate formula results.
// AI Prompts: Show C# code using Aspose.Cells to freeze panes at a specific cell and then recalculate all formulas before saving. | Explain why Workbook.CalculateFormula should be invoked after Worksheet.FreezePanes and list alternative ways to trigger a formula refresh in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, adds values and formulas, freezes the top two rows and left two columns at cell C3, then calls Workbook.CalculateFormula() to refresh all dependent calculations before saving.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data and formulas (so we have something to recalculate)
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["B1"].Formula = "=A1+A2";   // B1 = 30
        worksheet.Cells["C1"].Formula = "=B1*2";    // C1 = 60

        // Freeze panes at cell C3 (row index 2, column index 2) with 2 frozen rows and 2 frozen columns
        // Rule: Worksheet.FreezePanes(int, int, int, int)
        worksheet.FreezePanes(2, 2, 2, 2);

        // Recalculate all formulas after freezing panes
        // Rule: Workbook.CalculateFormula()
        workbook.CalculateFormula();

        // Save the workbook (lifecycle rule: save)
        workbook.Save("RecalculatedAfterFreeze.xlsx");
    }
}
