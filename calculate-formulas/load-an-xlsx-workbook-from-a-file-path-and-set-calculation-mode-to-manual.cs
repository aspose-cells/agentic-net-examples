using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class LoadAndSetManualCalculation
    {
        static void Main()
        {
            // Path to the existing XLSX file
            string filePath = "input.xlsx";

            // Load the workbook from the specified file path
            Workbook workbook = new Workbook(filePath);

            // Set the calculation mode to Manual
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // (Optional) Save the workbook to persist the setting
            // workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}