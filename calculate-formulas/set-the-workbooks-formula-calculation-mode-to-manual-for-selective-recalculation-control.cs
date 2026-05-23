using System;
using Aspose.Cells;

namespace AsposeCellsFormulaModeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data and a formula
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].Formula = "=A1+A2"; // Will not be calculated automatically

            // Set the calculation mode to Manual.
            // This tells Excel (and other applications) that the workbook should not recalculate
            // formulas automatically. You can trigger calculation manually via CalculateFormula().
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Optionally, calculate formulas now if you need the results before saving
            // workbook.CalculateFormula();

            // Save the workbook to an XLSX file
            workbook.Save("ManualCalculationMode.xlsx", SaveFormat.Xlsx);

            // Inform the user
            Console.WriteLine("Workbook saved with CalculationMode set to Manual.");
        }
    }
}