using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the current maximum row limit for shared formulas
        int currentLimit = workbook.Settings.MaxRowsOfSharedFormula;
        Console.WriteLine("Current MaxRowsOfSharedFormula: " + currentLimit);

        // Demonstrate the effect of a low limit
        workbook.Settings.MaxRowsOfSharedFormula = 100; // set a low limit
        Worksheet sheet1 = workbook.Worksheets[0];
        Cells cells1 = sheet1.Cells;

        // Attempt to set a shared formula that exceeds the limit (150 rows)
        cells1["B1"].SetSharedFormula("=A1", 150, 1);
        Console.WriteLine("Formula in B150 with low limit: " + cells1["B150"].Formula);

        // Increase the limit and set the same shared formula on a new sheet
        workbook.Settings.MaxRowsOfSharedFormula = 2000; // raise the limit
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        Cells cells2 = sheet2.Cells;
        cells2["B1"].SetSharedFormula("=A1", 150, 1);
        Console.WriteLine("Formula in B150 on Sheet2 with higher limit: " + cells2["B150"].Formula);

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("output.xlsx");
    }
}