using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data and formulas
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["B1"].Formula = "=A1+A2";   // B1 = 30
        worksheet.Cells["C1"].Formula = "=B1*2";    // C1 = 60

        // Freeze panes at cell C3 (row index 2, column index 2) with 2 rows and 2 columns frozen
        worksheet.FreezePanes(2, 2, 2, 2);

        // Recalculate all formulas so dependent cells reflect the frozen view
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("RecalcAfterFreeze.xlsx");
    }
}