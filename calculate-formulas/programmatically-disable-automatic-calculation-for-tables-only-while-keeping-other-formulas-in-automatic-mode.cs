using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Disable automatic calculation for tables only
        // Other formulas will still be calculated automatically when Excel opens the file
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

        // ----- Sample data (optional) -----
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);
        // Regular formula (outside any table)
        sheet.Cells["B1"].Formula = "=SUM(A1:A3)";

        // Create a simple table (ListObject) to illustrate table behavior
        // Table range: A5:B7 (3 rows, 2 columns)
        sheet.Cells["A5"].PutValue(1);
        sheet.Cells["A6"].PutValue(2);
        sheet.Cells["A7"].PutValue(3);
        sheet.Cells["B5"].Formula = "=A5*2";
        sheet.Cells["B6"].Formula = "=A6*2";
        sheet.Cells["B7"].Formula = "=A7*2";
        sheet.ListObjects.Add(4, 0, 6, 1, true); // adds table covering A5:B7

        // Save the workbook
        workbook.Save("TableCalcMode.xlsx");
    }
}