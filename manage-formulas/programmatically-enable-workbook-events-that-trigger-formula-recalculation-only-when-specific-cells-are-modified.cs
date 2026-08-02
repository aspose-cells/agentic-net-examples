// Title: Selective Formula Recalculation in Aspose.Cells (C#) – Trigger Calculation Only for Specified Cells
// Description: Demonstrates how to set Aspose.Cells to Manual calculation mode and recalculate formulas only when designated cells (e.g., A1, B1) are changed. A helper method updates a cell, checks a trigger list, and calls Workbook.CalculateFormula() selectively, improving performance for large workbooks.
// Keywords: Aspose.Cells manual calculation | selective formula recalculation | C# Aspose.Cells trigger cells | conditional workbook calculation | performance optimization Aspose.Cells | CalculateFormula on demand | Excel formula event simulation
// Common Searches: Aspose.Cells recalculate formulas only when certain cells change | C# manual calculation mode with custom triggers Aspose.Cells | how to avoid full workbook recalculation in Aspose.Cells | selective formula update Aspose.Cells .NET example | triggered calculation Aspose.Cells workbook
// Developer Intent: The developer wants to recalculate workbook formulas only after specific input cells are modified, avoiding unnecessary calculations.
// Use Cases: Large financial models where only key input cells should cause a full recalculation, reducing CPU load. | Data‑entry forms that update dependent results only after critical parameters are edited. | Custom event‑like handling in Aspose.Cells where changes to certain cells invoke CalculateFormula on demand.
// AI Prompts: Generate C# code using Aspose.Cells that recalculates formulas only when cells from a configurable list are edited. | Show how to extend the helper method to support multiple worksheets and dynamic trigger collections. | Refactor the example into a reusable class that mimics workbook events for selective formula calculation.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates how to set Aspose.Cells to Manual calculation mode and recalculate formulas only when designated cells (e.g., A1, B1) are changed. A helper method updates a cell, checks a trigger list, and calls Workbook.CalculateFormula() selectively, improving performance for large workbooks.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some initial data and a dependent formula
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(20);
        sheet.Cells["C1"].Formula = "=A1+B1"; // C1 depends on A1 and B1

        // Set calculation mode to Manual so formulas are not auto‑calculated
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Define the cells that should trigger a full recalculation when modified
        HashSet<string> triggerCells = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "A1",
            "B1"
        };

        // Helper method: set a cell's value and recalculate only if it is a trigger cell
        void SetCellValue(string address, object value)
        {
            sheet.Cells[address].PutValue(value);

            // If the modified cell is in the trigger list, recalculate the workbook
            if (triggerCells.Contains(address))
            {
                workbook.CalculateFormula();
            }
        }

        // Example modifications
        SetCellValue("A1", 30); // Triggers recalculation (C1 will be updated)
        SetCellValue("D1", 5);  // Does NOT trigger recalculation

        // Save the workbook (lifecycle rule: create → save)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
