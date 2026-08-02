using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook wb = new Workbook();

        // Access the default first worksheet and give it an initial name
        Worksheet sheet1 = wb.Worksheets[0];
        sheet1.Name = "Data";

        // Populate some data in the first worksheet
        sheet1.Cells["A1"].PutValue(10);
        sheet1.Cells["A2"].PutValue(20);

        // Add a second worksheet that contains a formula referencing the first sheet
        Worksheet sheet2 = wb.Worksheets.Add("Summary");
        // Formula uses the sheet name "Data"
        sheet2.Cells["B1"].Formula = "=Data!A1+Data!A2";

        // Initial calculation so the formula has a value
        wb.CalculateFormula();

        Console.WriteLine("Before rename, B1 value: " + sheet2.Cells["B1"].Value);

        // Rename the first worksheet
        sheet1.Name = "RenamedData";

        // Refresh all formulas after the rename operation
        wb.CalculateFormula();

        Console.WriteLine("After rename, B1 value: " + sheet2.Cells["B1"].Value);

        // Save the workbook (lifecycle: save)
        wb.Save("RenamedSheetDemo.xlsx");
    }
}