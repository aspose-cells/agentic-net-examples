// Title: Aspose.Cells C# – Set Workbook Calculation Mode to AutomaticExceptTable
// Description: Demonstrates how to configure Aspose.Cells Workbook.Settings.FormulaSettings.CalculationMode to CalcModeType.AutomaticExceptTable so that only table formulas recalculate automatically while all other formulas stay static, then saves the workbook.
// Keywords: Aspose.Cells calculation mode | AutomaticExceptTable C# | Aspose.Cells formula settings | disable automatic recalculation Aspose | Excel table formulas recalc | Aspose.Cells .NET example | Workbook.Settings.FormulaSettings | CalcModeType AutomaticExceptTable | Aspose.Cells GitHub sample | C# Excel automation
// Common Searches: Aspose.Cells set calculation mode AutomaticExceptTable | C# Aspose.Cells only recalculate table formulas | prevent non‑table formulas from auto‑calculating Aspose | how to enable AutomaticExceptTable in Aspose.Cells | save calculation mode in Excel file using Aspose.Cells
// Developer Intent: Configure a workbook so that only table formulas auto‑recalculate, leaving other formulas unchanged.
// Use Cases: Create an Excel template where static formulas stay fixed but table formulas stay up‑to‑date. | Export data with Aspose.Cells while preserving a custom calculation mode for downstream users. | Load an existing workbook, switch to AutomaticExceptTable to improve performance, and save without triggering full recalculation.
// AI Prompts: Write C# code with Aspose.Cells to set Workbook.Settings.FormulaSettings.CalculationMode to AutomaticExceptTable and describe its impact. | Show how to load an existing workbook, change its calculation mode to AutomaticExceptTable, and save it without forcing a full recalculation. | Explain how to query the current calculation mode of a workbook using Aspose.Cells .NET API.

using System;
using Aspose.Cells;

// Demonstrates how to configure Aspose.Cells Workbook.Settings.FormulaSettings.CalculationMode to CalcModeType.AutomaticExceptTable so that only table formulas recalculate automatically while all other formulas stay static, then saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set calculation mode to AutomaticExceptTable.
        // Table formulas will be recalculated automatically,
        // while other formulas stay static unless calculated manually.
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

        // Save the workbook (the calculation mode is stored in the file)
        workbook.Save("AutomaticExceptTable.xlsx");
    }
}
