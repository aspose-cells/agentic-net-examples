// Title: C# – Load Excel workbook, set manual calculation mode, and save with Aspose.Cells
// Description: Demonstrates how to open an existing .xlsx file using Aspose.Cells for .NET, change the workbook's formula calculation mode to Manual, and write the result to a new file.
// Keywords: Aspose.Cells manual calculation mode | C# load Excel workbook | disable automatic formula recalculation | Aspose.Cells save workbook | set CalcModeType.Manual | Excel performance optimization Aspose
// Common Searches: Aspose.Cells set calculation mode to manual C# | how to prevent formula recalculation when loading Excel with Aspose | save workbook after changing formula settings Aspose.Cells | manual calculation mode example Aspose.Cells .NET | improve Excel load speed Aspose.Cells manual mode
// Developer Intent: Change a workbook’s formula calculation setting to Manual and persist the change to a new file.
// Use Cases: Speed up loading of large spreadsheets by suppressing immediate formula evaluation. | Prepare a workbook for bulk data updates without triggering recalculation after each change. | Export a workbook to another system while keeping formulas unevaluated until they are needed.
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, sets the calculation mode to Manual, and saves it under a different name. | Explain why and how to disable automatic formula calculation in Aspose.Cells, including performance benefits. | Provide a step‑by‑step tutorial for changing a workbook’s calculation mode to Manual and saving it using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates how to open an existing .xlsx file using Aspose.Cells for .NET, change the workbook's formula calculation mode to Manual, and write the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook from disk using the string constructor
            Workbook workbook = new Workbook("input.xlsx");

            // Set the calculation mode to Manual
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Save the workbook back to disk
            workbook.Save("output.xlsx");
        }
    }
}
