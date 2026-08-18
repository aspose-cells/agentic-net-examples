// Title: Set Aspose.Cells CalculationMode to AutomaticExceptTable (C#)
// Description: Demonstrates how to create a workbook, switch its FormulaSettings.CalculationMode to CalcModeType.AutomaticExceptTable so that only non‑table formulas recalculate automatically, verify the setting, and save the file as XLSX.
// Keywords: Aspose.Cells calculation mode | AutomaticExceptTable C# | CalcModeType AutomaticExceptTable | exclude table formulas | .NET Excel recalculation | Aspose.Cells FormulaSettings
// Common Searches: Aspose.Cells set AutomaticExceptTable mode | C# disable automatic recalculation for Excel tables | How to use CalcModeType.AutomaticExceptTable in .NET | Aspose.Cells formula calculation options | Prevent table formulas from auto‑updating with Aspose
// Developer Intent: Configure a workbook so that Excel automatically recalculates regular formulas while leaving table‑based formulas untouched.
// Use Cases: Create a new workbook and apply AutomaticExceptTable to improve performance when many table formulas exist. | Programmatically confirm the active calculation mode before exporting the file. | Switch between Automatic, Manual, and AutomaticExceptTable depending on the processing scenario.
// AI Prompts: Generate C# code that sets Aspose.Cells Workbook.CalculationMode to AutomaticExceptTable and saves the workbook. | Explain the differences between Automatic, Manual, and AutomaticExceptTable calculation modes in Aspose.Cells. | Provide a step‑by‑step guide for toggling calculation modes in a .NET application using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCalculationModeExample
{
    // Demonstrates how to create a workbook, switch its FormulaSettings.CalculationMode to CalcModeType.AutomaticExceptTable so that only non‑table formulas recalculate automatically, verify the setting, and save the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided creation rule)
            Workbook workbook = new Workbook();

            // Set the calculation mode to AutomaticExceptTable.
            // This mode tells Excel to recalculate formulas automatically
            // except those that belong to Excel tables.
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

            // Optional: display the current mode to verify
            Console.WriteLine("Current CalculationMode: " + workbook.Settings.FormulaSettings.CalculationMode);

            // Save the workbook (uses the provided save rule)
            workbook.Save("CalculationMode_AutomaticExceptTable.xlsx", SaveFormat.Xlsx);
        }
    }
}
