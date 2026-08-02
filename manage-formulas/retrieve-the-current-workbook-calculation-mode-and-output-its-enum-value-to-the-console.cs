// Title: Get Workbook Calculation Mode (CalcModeType) with Aspose.Cells for .NET
// Description: Demonstrates how to read the current calculation mode of a workbook via Settings.FormulaSettings.CalculationMode and write the CalcModeType enum value to the console.
// Keywords: Aspose.Cells calculation mode | CalcModeType | Workbook Settings FormulaSettings | C# read calculation mode | Aspose.Cells .NET get CalcModeType | retrieve workbook calculation mode | FormulaSettings.CalculationMode API
// Common Searches: Aspose.Cells get calculation mode | How to read CalcModeType in C# | Workbook Settings FormulaSettings CalculationMode example | Aspose.Cells current calculation mode | Retrieve workbook calculation mode programmatically
// Developer Intent: Read the workbook's current calculation mode and display the enum value.
// Use Cases: Log the calculation mode when loading a template to confirm Automatic or Manual setting. | Determine if a switch to Manual mode is needed before performing bulk cell updates. | Debug formula evaluation by outputting the CalcModeType after opening a workbook.
// AI Prompts: Generate C# code using Aspose.Cells that reads the workbook's CalculationMode and prints the enum. | Show how to compare the retrieved CalcModeType with Automatic, Manual, and AutomaticExceptTables and modify it if required. | Explain how to log the current calculation mode for troubleshooting formula evaluation in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to read the current calculation mode of a workbook via Settings.FormulaSettings.CalculationMode and write the CalcModeType enum value to the console.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Get the current calculation mode from the workbook's formula settings
        CalcModeType currentMode = workbook.Settings.FormulaSettings.CalculationMode;

        // Output the enum value to the console
        Console.WriteLine("Current Calculation Mode: " + currentMode);
    }
}
