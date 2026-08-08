// Title: Aspose.Cells .NET – Switch to Manual Calculation, Bulk‑Update Cells, Then Recalculate Once
// Description: Learn how to set Workbook.Settings.FormulaSettings.CalculationMode to Manual, insert thousands of values without triggering recalculation, add formulas, and finally call Workbook.CalculateFormula a single time for optimal performance.
// Keywords: Aspose.Cells manual calculation mode | Workbook.Settings.FormulaSettings | CalculateFormula after bulk update | disable automatic formula evaluation .NET | bulk cell insert Aspose.Cells | performance optimization Aspose.Cells | C# Aspose.Cells example | Workbook.CalculateFormula | large data import Aspose.Cells
// Common Searches: how to disable automatic calculation in Aspose.Cells .NET | bulk insert rows and recalculate formulas Aspose.Cells | manual calculation mode performance Aspose.Cells | set CalculationMode to Manual Aspose.Cells C# | calculate workbook once after data load Aspose.Cells
// Developer Intent: Defer formula evaluation while loading large data sets, then run a single calculation pass to improve speed.
// Use Cases: Import thousands of records into a worksheet without per‑row recalculation. | Generate financial reports where all formulas are evaluated only after data entry completes. | Create data‑intensive spreadsheets (e.g., sensor logs) and compute aggregates efficiently.
// AI Prompts: Show me C# code to set Workbook.Settings.FormulaSettings.CalculationMode to Manual, bulk‑fill cells, and then call Workbook.CalculateFormula once. | Explain the performance impact of manual calculation mode when inserting 10,000 rows with Aspose.Cells. | Provide a step‑by‑step guide for loading large CSV data into an Aspose.Cells workbook while keeping calculation manual and recalculating at the end.

using System;
using Aspose.Cells;

namespace AsposeCellsBulkUpdateExample
{
    // Learn how to set Workbook.Settings.FormulaSettings.CalculationMode to Manual, insert thousands of values without triggering recalculation, add formulas, and finally call Workbook.CalculateFormula a single time for optimal performance.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Switch calculation mode to Manual to defer formula evaluation
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Perform bulk updates – fill column A with numbers 1..1000
            for (int i = 0; i < 1000; i++)
            {
                cells[i, 0].PutValue(i + 1); // Row i, Column 0 (A)
            }

            // Add a formula that sums the range we just filled
            cells["B1"].Formula = "=SUM(A1:A1000)";

            // After all updates, trigger a single calculation pass
            workbook.CalculateFormula();

            // Optional: display the calculated sum in console
            Console.WriteLine("Sum of A1:A1000 = " + cells["B1"].Value);

            // Save the workbook (using the standard save method)
            workbook.Save("BulkUpdateResult.xlsx");
        }
    }
}
