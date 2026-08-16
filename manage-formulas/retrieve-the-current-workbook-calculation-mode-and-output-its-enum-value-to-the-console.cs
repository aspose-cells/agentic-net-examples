// Title: Get Workbook Calculation Mode (CalcModeType) with Aspose.Cells for .NET
// Description: Demonstrates how to read the current CalcModeType enum from a workbook's FormulaSettings and print it to the console using Aspose.Cells for .NET.
// Keywords: Aspose.Cells calculation mode | CalcModeType enum | C# workbook formula settings | read workbook calculation mode | Aspose.Cells console output
// Common Searches: Aspose.Cells get current calculation mode C# | How to read CalcModeType from workbook | Display workbook formula calculation mode | Aspose.Cells formula settings example | C# retrieve workbook calculation mode
// Developer Intent: Read the workbook's current calculation mode and output the enum value.
// Use Cases: Confirm that a workbook uses Automatic calculation before bulk updates. | Log the calculation mode of imported workbooks for troubleshooting formula evaluation. | Switch to Manual mode programmatically after checking the current setting to improve performance.
// AI Prompts: Generate C# code with Aspose.Cells that prints the workbook's CalcModeType to the console. | Create a method that returns the current calculation mode of a given Workbook and logs it. | Show how to detect Automatic calculation mode and change it to Manual using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCalculationModeDemo
{
    // Demonstrates how to read the current CalcModeType enum from a workbook's FormulaSettings and print it to the console using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default settings)
            Workbook workbook = new Workbook();

            // Retrieve the current calculation mode from the workbook's formula settings
            CalcModeType currentMode = workbook.Settings.FormulaSettings.CalculationMode;

            // Output the enum value to the console
            Console.WriteLine("Current Calculation Mode: " + currentMode);
        }
    }
}
