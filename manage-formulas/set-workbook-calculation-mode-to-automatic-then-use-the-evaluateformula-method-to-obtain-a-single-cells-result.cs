// Title: Aspose.Cells .NET: Set Automatic Calculation Mode and Evaluate a Formula Directly
// Description: This example demonstrates how to configure a workbook for automatic recalculation using Aspose.Cells, insert sample values, and obtain the result of an Excel formula on‑the‑fly with `worksheet.CalculateFormula`. The result is printed to the console and the workbook can be saved if needed.
// Keywords: Aspose.Cells automatic calculation | C# evaluate Excel formula programmatically | worksheet.CalculateFormula example | set calculation mode Aspose.Cells | .NET Excel formula evaluation | Aspose.Cells formula settings
// Common Searches: Aspose.Cells set calculation mode to automatic | How to evaluate a formula without placing it in a cell using Aspose.Cells | C# calculate =A1+B1 with Aspose.Cells | Automatic workbook recalculation Aspose.Cells .NET | Evaluate Excel formula programmatically in C#
// Developer Intent: Configure a workbook for automatic calculation and retrieve a formula's result without writing the formula to a worksheet cell.
// Use Cases: Automatically recalculate dependent formulas after updating cell values. | Fetch a quick calculation result for user input before persisting the workbook. | Perform ad‑hoc formula evaluations in server‑side processing or API services.
// AI Prompts: Show how to switch the calculation mode to Manual and trigger a full workbook recalculation with Aspose.Cells. | Provide a C# snippet that uses worksheet.CalculateFormula to evaluate a nested formula referencing multiple worksheets. | Explain how to capture calculation errors when using CalculateFormula in Aspose.Cells.

using System;
using Aspose.Cells;

// This example demonstrates how to configure a workbook for automatic recalculation using Aspose.Cells, insert sample values, and obtain the result of an Excel formula on‑the‑fly with `worksheet.CalculateFormula`. The result is printed to the console and the workbook can be saved if needed.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set calculation mode to Automatic
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Populate some sample data
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["B1"].PutValue(20);

        // Evaluate a formula directly (no need to place it in a cell)
        object result = worksheet.CalculateFormula("=A1+B1");

        // Display the evaluated result
        Console.WriteLine("Result of =A1+B1: " + result);

        // Save the workbook (optional)
        workbook.Save("CalculatedResult.xlsx");
    }
}
