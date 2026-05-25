using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Disable shared formulas for the entire workbook
        workbook.Settings.MaxRowsOfSharedFormula = 0;

        // Example data and a regular formula (not shared)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["B1"].Formula = "=A1+A2";

        // Calculate formulas
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("DisabledSharedFormulas.xlsx");
    }
}