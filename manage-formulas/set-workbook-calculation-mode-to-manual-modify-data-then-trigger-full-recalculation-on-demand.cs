using System;
using Aspose.Cells;

namespace AsposeCellsCalcModeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add initial data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Add a formula that sums the three values
            cells["B1"].Formula = "=SUM(A1:A3)";

            // Access formula settings and set calculation mode to Manual
            FormulaSettings formulaSettings = workbook.Settings.FormulaSettings;
            formulaSettings.CalculationMode = CalcModeType.Manual;

            // Ensure a full recalculation is performed when we trigger it
            formulaSettings.ForceFullCalculation = true;

            // Save the workbook before any calculation (optional)
            workbook.Save("BeforeCalculation.xlsx");

            // Modify data after setting manual mode
            cells["A2"].PutValue(50); // Change a value that affects the sum

            // Trigger full recalculation on demand
            workbook.CalculateFormula();

            // Verify the result (optional console output)
            Console.WriteLine("Recalculated B1 value: " + cells["B1"].Value);

            // Save the workbook after calculation
            workbook.Save("AfterCalculation.xlsx");
        }
    }
}