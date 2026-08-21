// Title: Set Automatic Calculation Mode After Merging Cells with Aspose.Cells for .NET (C#)
// Description: Shows how to merge a range, switch the workbook's calculation mode to Automatic, trigger an immediate recalculation, and save the file so formulas that reference the merged cells update instantly.
// Keywords: Aspose.Cells | C# | automatic calculation mode | merge cells | recalculate formulas | FormulaSettings | CalcModeType.Automatic | Workbook.CalculateFormula | Excel merging | refresh dependent formulas
// Common Searches: Aspose.Cells set calculation mode to automatic after merge | C# merge cells and recalculate formulas Aspose.Cells | How to refresh formulas after merging cells in Aspose.Cells | Automatic formula calculation in Aspose.Cells .NET | Force recalculation after cell merge Aspose.Cells
// Developer Intent: Enable automatic formula recalculation immediately after merging cells so dependent calculations stay up‑to‑date without manual intervention.
// Use Cases: Merging header rows in a financial report while keeping total formulas current. | Consolidating data rows in a template and ensuring summary calculations update on the fly. | Building a dashboard where merged cells trigger instant updates of charts and KPI formulas.
// AI Prompts: Generate C# code using Aspose.Cells to merge a range, set the workbook's calculation mode to Automatic, and recalculate all formulas. | Demonstrate switching the calculation mode to Manual, performing several merges, then invoking a single CalculateFormula call with Aspose.Cells. | Explain how to programmatically verify that formulas referencing merged cells return the expected values after the merge.

using System;
using Aspose.Cells;

// Shows how to merge a range, switch the workbook's calculation mode to Automatic, trigger an immediate recalculation, and save the file so formulas that reference the merged cells update instantly.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);

        // Add a formula that depends on the cells to be merged
        sheet.Cells["B1"].Formula = "=SUM(A1:A2)";

        // Merge cells A1:A2 (rows 0-1, column 0)
        sheet.Cells.Merge(0, 0, 2, 1);

        // Set calculation mode to Automatic so dependent formulas recalculate instantly
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Force immediate recalculation of all formulas
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("MergedWorkbook.xlsx");
    }
}
