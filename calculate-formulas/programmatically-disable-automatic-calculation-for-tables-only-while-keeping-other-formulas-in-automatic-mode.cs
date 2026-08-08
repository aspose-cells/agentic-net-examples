// Title: Disable automatic calculation for tables only while keeping other formulas automatic with Aspose.Cells for .NET
// Description: Learn how to set Aspose.Cells' CalculationMode to AutomaticExceptTable in C#. The workbook (new or existing) will recalculate all non‑table formulas automatically, while table formulas are excluded until manually refreshed. Finally, the workbook is saved with the new setting.
// Keywords: Aspose.Cells AutomaticExceptTable | C# calculation mode tables | disable table auto‑recalc Aspose | Aspose.Cells formula settings .NET | Excel workbook calculation mode | AutomaticExceptTable example
// Common Searches: Aspose.Cells set CalculationMode to AutomaticExceptTable | disable automatic calculation for tables only C# | keep non‑table formulas auto‑calculating Aspose.Cells | how to exclude tables from auto‑recalc in Aspose.Cells | programmatically change calculation mode Aspose.Cells
// Developer Intent: Configure a workbook so that only table formulas are excluded from automatic recalculation, while all other formulas remain in automatic mode.
// Use Cases: Large data‑driven reports where table formulas are refreshed manually to improve performance. | Batch processing of spreadsheets where table recalculation is costly, using AutomaticExceptTable to speed up updates. | Creating a template workbook that automatically updates standard formulas but leaves table calculations under user control.
// AI Prompts: Generate C# code to revert a workbook from AutomaticExceptTable back to full Automatic calculation using Aspose.Cells. | Show how to trigger manual recalculation of only the tables after setting CalculationMode to AutomaticExceptTable. | Compare performance of AutomaticExceptTable versus full Automatic mode for workbooks with thousands of rows.

using System;
using Aspose.Cells;

// Learn how to set Aspose.Cells' CalculationMode to AutomaticExceptTable in C#. The workbook (new or existing) will recalculate all non‑table formulas automatically, while table formulas are excluded until manually refreshed. Finally, the workbook is saved with the new setting.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // new Workbook("input.xlsx") to load

        // Disable automatic calculation for tables only.
        // Other formulas will still be calculated automatically.
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

        // Save the workbook with the new calculation mode setting.
        workbook.Save("output.xlsx");
    }
}
