// Title: Set Aspose.Cells workbook calculation mode to AutomaticExceptTable in C# to exclude table formulas from auto‑recalculation
// AI Prompts: Write C# code that creates an Aspose.Cells workbook and sets FormulaSettings.CalculationMode to AutomaticExceptTable so that table formulas are not recalculated automatically. | Show how to read back and display the current calculation mode after changing it, then save the workbook as an .xlsx file.
// Common Searches: aspnet cells c# set calculation mode automaticexcepttable | how to prevent table formulas from auto updating in Aspose.Cells | C# Aspose.Cells change workbook formula calculation mode programmatically | example of using CalcModeType.AutomaticExceptTable with Aspose.Cells | skip table formula recalculation Aspose.Cells workbook settings
// Tags: Aspose.Cells calculation mode without table formulas | C# set workbook formula calculation mode Aspose.Cells | exclude table formulas from auto recalculation Aspose.Cells | configure CalcModeType in Aspose.Cells | Aspose.Cells formula recalculation settings

using System;
using Aspose.Cells;

// Demonstrates creating a new Aspose.Cells workbook in C#, configuring its FormulaSettings.CalculationMode to AutomaticExceptTable to skip table formulas during automatic recalculation, printing the mode, and saving the workbook to an .xlsx file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set calculation mode to AutomaticExceptTable to exclude table formulas from automatic updates
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

        // Display the current calculation mode
        Console.WriteLine("Calculation Mode: " + workbook.Settings.FormulaSettings.CalculationMode);

        // Save the workbook (optional)
        workbook.Save("CalculationModeDemo.xlsx");
    }
}
