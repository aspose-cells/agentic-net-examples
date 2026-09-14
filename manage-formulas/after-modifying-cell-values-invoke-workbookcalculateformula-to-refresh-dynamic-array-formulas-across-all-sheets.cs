// Title: Recalculate all formulas, including dynamic arrays, after updating cells with Aspose.Cells for .NET
// AI Prompts: Load a workbook, change specific cell values, then call Workbook.CalculateFormula to refresh every formula in all worksheets using Aspose.Cells for .NET. | Demonstrate how to trigger a full workbook recalculation after modifying cells A1 and B2 with Aspose.Cells. | Show the steps to update cell data and invoke CalculateFormula to ensure dynamic array results are updated across the entire file.
// Common Searches: Aspose.Cells how to recalculate dynamic array formulas after editing cells | C# workbook.CalculateFormula across all worksheets example | Refresh formulas in an Excel file after changing cell values using Aspose.Cells .NET | Trigger full formula evaluation in Aspose.Cells after updating cell A1 and B2 | CalculateFormula method to update dynamic arrays in Aspose.Cells workbook
// Tags: Workbook.CalculateFormula for full workbook recalculation | dynamic array formula refresh Aspose.Cells | update cell values then recalc Excel workbook .NET | recalculate formulas across multiple worksheets Aspose.Cells | C# Aspose.Cells modify cells and recalculate

using Aspose.Cells;

// The example loads 'input.xlsx', updates cells A1 and B2, calls Workbook.CalculateFormula to recompute all formulas—including dynamic arrays—across every worksheet, and saves the modified workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Example: modify some cell values
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);   // set A1 to 10
        sheet.Cells["B2"].PutValue(20);   // set B2 to 20

        // Recalculate all formulas (including dynamic array formulas) across all sheets
        workbook.CalculateFormula();

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}
