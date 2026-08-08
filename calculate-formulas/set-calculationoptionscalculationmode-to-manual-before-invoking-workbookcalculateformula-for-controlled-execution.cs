// Title: Aspose.Cells .NET: Set Workbook CalculationMode to Manual and Evaluate Formulas with CalculationOptions
// Description: Demonstrates how to switch a workbook to manual calculation mode, add formulas, configure CalculationOptions, and explicitly trigger formula evaluation using Workbook.CalculateFormula in C#. Includes performance tips and saving the result.
// Keywords: Aspose.Cells | .NET | C# | CalculationMode Manual | CalculateFormula | CalculationOptions | manual formula evaluation | disable automatic recalculation | spreadsheet performance | controlled calculation
// Common Searches: Aspose.Cells set calculation mode manual | Workbook.CalculateFormula with options C# | manual formula calculation example Aspose.Cells | how to prevent auto recalculation Aspose.Cells | performance optimization Aspose.Cells manual mode
// Developer Intent: Configure a workbook for manual calculation and invoke CalculateFormula with custom options to control when and how formulas are evaluated.
// Use Cases: Compute formulas only after bulk data updates to avoid repeated recalculations. | Run intensive spreadsheet calculations in a background thread without UI blocking. | Apply specific CalculationOptions (e.g., precision, error handling) before evaluating formulas.
// AI Prompts: Show C# code to set Aspose.Cells Workbook calculation mode to Manual and calculate formulas with CalculationOptions. | Explain when to use manual vs. automatic calculation in Aspose.Cells and how to switch modes. | Provide examples of customizing CalculationOptions before calling Workbook.CalculateFormula.

using System;
using Aspose.Cells;

namespace AsposeCellsCalculationModeDemo
{
    // Demonstrates how to switch a workbook to manual calculation mode, add formulas, configure CalculationOptions, and explicitly trigger formula evaluation using Workbook.CalculateFormula in C#. Includes performance tips and saving the result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample data and formulas
            cells["A1"].PutValue(5);
            cells["B1"].Formula = "=A1*2";
            cells["C1"].Formula = "=B1+10";

            // Set calculation mode to Manual (controlled execution)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Create calculation options (optional customizations can be added)
            CalculationOptions calcOptions = new CalculationOptions();

            // Perform calculation using the specified options
            workbook.CalculateFormula(calcOptions);

            // Output results to console
            Console.WriteLine("A1 value: " + cells["A1"].IntValue);
            Console.WriteLine("B1 calculated value: " + cells["B1"].IntValue);
            Console.WriteLine("C1 calculated value: " + cells["C1"].IntValue);

            // Save the workbook (optional)
            workbook.Save("ManualCalculationDemo.xlsx");
        }
    }
}
