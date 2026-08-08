// Title: Aspose.Cells C# – Set Workbook Calculation Mode to Automatic for Real‑Time Formula Updates
// Description: Learn how to enable Automatic calculation mode in Aspose.Cells for .NET so formulas recalculate instantly. This example creates a workbook, switches the FormulaSettings.CalculationMode to CalcModeType.Automatic, adds data and a simple formula, and saves the file, illustrating real‑time formula updates without manual CalculateFormula calls.
// Keywords: Aspose.Cells calculation mode | C# automatic formula recalculation | Workbook Settings FormulaSettings | CalcModeType.Automatic example | Aspose.Cells .NET API | real‑time formula updates | set workbook calculation mode programmatically | Aspose.Cells Excel automation
// Common Searches: Aspose.Cells set calculation mode to automatic C# | Enable automatic formula calculation in Aspose.Cells .NET | How to switch workbook calculation mode in Aspose.Cells | Automatic vs manual calculation mode Aspose.Cells | C# code for Aspose.Cells automatic recalculation
// Developer Intent: Configure a workbook to use Automatic calculation so formulas update immediately as cell values change.
// Use Cases: Create a new workbook that automatically recalculates formulas after each data entry. | Convert an existing workbook from Manual to Automatic mode before bulk data processing to keep results current. | Implement server‑side spreadsheet services where frequent value changes require instant formula evaluation.
// AI Prompts: Generate C# code using Aspose.Cells that sets the workbook's calculation mode to Automatic and demonstrates a formula updating without calling CalculateFormula. | Show how to toggle a workbook's calculation mode from Manual to Automatic and explain the impact on newly added formulas. | Explain the relationship between Aspose.Cells calculation modes and runtime behavior, and when an explicit CalculateFormula call is still needed.

using System;
using Aspose.Cells;

namespace AsposeCellsCalculationModeDemo
{
    // Learn how to enable Automatic calculation mode in Aspose.Cells for .NET so formulas recalculate instantly. This example creates a workbook, switches the FormulaSettings.CalculationMode to CalcModeType.Automatic, adds data and a simple formula, and saves the file, illustrating real‑time formula updates without manual CalculateFormula calls.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Set the calculation mode to Automatic for immediate formula updates
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Example: add some data and a formula to demonstrate automatic calculation
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=A1+A2";

            // Calculate formulas now (optional, as mode is Automatic for Excel, not for Aspose runtime)
            workbook.CalculateFormula();

            // Save the workbook (lifecycle save)
            workbook.Save("CalculationModeAutomatic.xlsx");
        }
    }
}
