using System;
using Aspose.Cells;

class DeleteColumnAndRecalculate
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add sample data
        cells["A1"].PutValue(10);
        cells["B1"].PutValue(20);
        cells["C1"].PutValue(30);

        // Formula that references the three columns
        cells["D1"].Formula = "=SUM(A1:C1)";

        // Delete column B (index 1) and update references in formulas
        cells.DeleteColumn(1, true);

        // Recalculate formulas after the column deletion
        workbook.CalculateFormula();

        // Save the modified workbook
        workbook.Save("AfterDeleteColumn.xlsx");
    }
}