using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set calculation mode to AutomaticExceptTable
            // This mode excludes table formulas from automatic recalculation.
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

            // Output the current setting to verify
            Console.WriteLine("Calculation Mode: " + workbook.Settings.FormulaSettings.CalculationMode);

            // Save the workbook (the file can be opened in Excel to see the setting)
            workbook.Save("CalculationMode_AutomaticExceptTable.xlsx");
        }
    }
}