using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // ---------- Worksheet 1 ----------
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        // Populate some data in Sheet1!A1:A3
        sheet1.Cells["A1"].PutValue(1);
        sheet1.Cells["A2"].PutValue(2);
        sheet1.Cells["A3"].PutValue(3);

        // ---------- Worksheet 2 ----------
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        // Populate some data in Sheet2!B1:B3
        sheet2.Cells["B1"].PutValue(10);
        sheet2.Cells["B2"].PutValue(20);
        sheet2.Cells["B3"].PutValue(30);

        // ---------- Worksheet 3 (where the formula will be used) ----------
        Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

        // Create a named range that spans both Sheet1 and Sheet2
        // The RefersTo string uses a union of two areas separated by a comma
        int nameIndex = workbook.Worksheets.Names.Add("MultiRange");
        Name multiRange = workbook.Worksheets.Names[nameIndex];
        multiRange.RefersTo = "=Sheet1!$A$1:$A$3,Sheet2!$B$1:$B$3";

        // Use the named range in a cross‑sheet formula on Sheet3
        sheet3.Cells["C1"].Formula = "=SUM(MultiRange)";

        // Calculate all formulas so that C1 contains the result
        workbook.CalculateFormula();

        // Display the calculated result
        Console.WriteLine("Sum of MultiRange = " + sheet3.Cells["C1"].Value);

        // Save the workbook
        workbook.Save("MultiSheetNamedRange.xlsx");
    }
}