// Title: Set Workbook CalculationMode to Automatic for all worksheets after data import – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to import data into each worksheet of a new Workbook, then enforce the workbook‑level FormulaSettings.CalculationMode = Automatic by iterating over the sheets, optionally recalculate formulas, and save the file as Output.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | CalculationMode Automatic | FormulaSettings | iterate worksheets | recalculate formulas | data import workbook | set automatic calculation | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set calculation mode automatic | C# loop through worksheets set formula settings | Enable automatic formula recalculation after data import Aspose.Cells | How to force workbook calculation mode in Aspose.Cells .NET | Calculate formulas programmatically Aspose.Cells
// Developer Intent: Apply the Automatic calculation mode to the workbook after populating data in every worksheet.
// Use Cases: Populate sales data across multiple sheets and ensure all dependent formulas refresh automatically before exporting. | Create a multi‑sheet financial model where each sheet contains formulas that must recalculate after batch data entry. | Automate report generation pipelines that require the workbook to be in Automatic mode to guarantee up‑to‑date calculations.
// AI Prompts: Generate C# code with Aspose.Cells that sets Workbook.Settings.FormulaSettings.CalculationMode to Automatic after filling cells in each worksheet. | Explain why CalculationMode is a workbook‑level property and how iterating over worksheets still achieves the desired effect. | Show how to invoke workbook.CalculateFormula() after changing the calculation mode in an Aspose.Cells .NET application.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates how to import data into each worksheet of a new Workbook, then enforce the workbook‑level FormulaSettings.CalculationMode = Automatic by iterating over the sheets, optionally recalculate formulas, and save the file as Output.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // ----- Data import simulation -----
            // For demonstration, add sample data to each worksheet.
            // In real scenarios, replace this block with actual data import logic.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add some values to the first worksheet
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(15);
            }

            // ----- Set calculation mode to Automatic for all worksheets -----
            // Although CalculationMode is a workbook‑level setting, we iterate over worksheets
            // as requested and set the mode repeatedly (the final value will be Automatic).
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
            }

            // Optionally, calculate formulas now if needed
            workbook.CalculateFormula();

            // Save the workbook to a file
            workbook.Save("Output.xlsx", SaveFormat.Xlsx);
        }
    }
}
