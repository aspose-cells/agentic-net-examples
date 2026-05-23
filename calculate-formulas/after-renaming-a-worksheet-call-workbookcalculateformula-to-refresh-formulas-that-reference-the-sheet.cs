using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet and give it an initial name
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Original";

        // Populate some data and a simple formula in the first sheet
        sheet1.Cells["A1"].PutValue(10);
        sheet1.Cells["A2"].PutValue(20);
        sheet1.Cells["B1"].Formula = "=SUM(A1:A2)";

        // Add a second worksheet that references the first sheet's formula
        Worksheet sheet2 = workbook.Worksheets.Add("Summary");
        sheet2.Cells["A1"].Formula = "=Original!B1";

        // Calculate all formulas so that initial values are available
        workbook.CalculateFormula();

        Console.WriteLine("Before rename, Summary!A1 = " + sheet2.Cells["A1"].StringValue);

        // Rename the first worksheet
        sheet1.Name = "Renamed";

        // Refresh formulas after the rename operation
        workbook.CalculateFormula();

        Console.WriteLine("After rename, Summary!A1 = " + sheet2.Cells["A1"].StringValue);

        // Save the workbook (saving rule)
        workbook.Save("RenamedSheetDemo.xlsx");
    }
}