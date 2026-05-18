using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class SetCalculationModeAfterMerge
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Set a formula that depends on the range A1:A3
            cells["B1"].Formula = "=SUM(A1:A3)";

            // Merge cells A1:A3 (merge rule)
            sheet.Cells.Merge(0, 0, 3, 1); // rows 0-2, columns 0-0 (A1:A3)

            // After merging, set calculation mode to Automatic to refresh dependent formulas instantly
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Force calculation of all formulas (calculate rule)
            workbook.CalculateFormula();

            // Save the workbook (save rule)
            workbook.Save("MergedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}