using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Set calculation mode to AutomaticExceptTable
            // This mode tells Excel to recalculate formulas automatically
            // except those that belong to Excel tables.
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

            // Optional: display the current mode to verify
            Console.WriteLine("Current CalculationMode: " + workbook.Settings.FormulaSettings.CalculationMode);

            // Save the workbook (the calculation mode is stored in the file)
            workbook.Save("CalculationMode_AutomaticExceptTable.xlsx");
        }
    }
}