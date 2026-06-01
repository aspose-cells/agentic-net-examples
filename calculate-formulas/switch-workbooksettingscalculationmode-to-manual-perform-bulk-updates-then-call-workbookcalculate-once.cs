using System;
using Aspose.Cells;

namespace BulkUpdateManualCalculation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: use provided creation method)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Switch calculation mode to Manual (required for bulk updates)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Perform bulk updates – for demonstration, fill a 1000‑row table
            for (int row = 0; row < 1000; row++)
            {
                // Column A: sequential numbers
                cells[row, 0].PutValue(row + 1);

                // Column B: random values
                cells[row, 1].PutValue(new Random().NextDouble() * 100);
            }

            // Add a formula that sums column B
            cells[1000, 0].Formula = $"=SUM(B1:B{1000})";

            // After all updates, calculate the workbook once
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: use provided save method)
            workbook.Save("BulkUpdateManualCalc.xlsx");
        }
    }
}