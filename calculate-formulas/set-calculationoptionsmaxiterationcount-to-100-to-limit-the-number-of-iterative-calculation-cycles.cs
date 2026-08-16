// Title: Set MaxIteration to 100 for iterative formula calculation in Aspose.Cells (.NET)
// Description: Shows how to enable iterative calculation, cap the number of iteration cycles at 100 with Workbook.Settings.FormulaSettings.MaxIteration, and persist the setting by saving the workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET | FormulaSettings | EnableIterativeCalculation | MaxIteration | iteration limit | circular reference handling | iterative calculation | limit iterations | workbook performance
// Common Searches: Aspose.Cells set MaxIteration to 100 | enable iterative calculation Aspose.Cells C# | limit circular reference iterations .NET | FormulaSettings MaxIteration example | how to cap iterative formula cycles in Aspose.Cells
// Developer Intent: Configure a workbook to perform no more than 100 iterative calculation cycles.
// Use Cases: Prevent endless loops when formulas contain circular references. | Improve calculation speed for large sheets that use iterative formulas. | Ensure deterministic results by fixing the maximum number of recalculation passes.
// AI Prompts: Provide C# code that enables iterative calculation and sets MaxIteration to 100 with Aspose.Cells. | Explain the effect of MaxIteration on circular reference resolution and workbook performance. | Show how to verify that the iteration limit is saved in the resulting Excel file.

using System;
using Aspose.Cells;

// Shows how to enable iterative calculation, cap the number of iteration cycles at 100 with Workbook.Settings.FormulaSettings.MaxIteration, and persist the setting by saving the workbook using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable iterative calculation to allow circular references to be resolved
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

        // Limit the number of iterative calculation cycles to 100
        workbook.Settings.FormulaSettings.MaxIteration = 100;

        // Perform calculation (optional, demonstrates that settings are applied)
        workbook.CalculateFormula();

        // Save the workbook to verify the settings are persisted
        workbook.Save("IterativeCalculation.xlsx");
    }
}
