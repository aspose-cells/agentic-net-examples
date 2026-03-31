using System;
using Aspose.Cells;

class ShowFormulasDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a formula in cell A1
        sheet.Cells["A1"].Formula = "=1+2+3";

        // Initially display the calculated result
        sheet.ShowFormulas = false;
        Console.WriteLine("Calculated value displayed: " + sheet.Cells["A1"].StringValue);

        // Enable formula display
        sheet.ShowFormulas = true;
        Console.WriteLine("Formula displayed: " + sheet.Cells["A1"].StringValue);

        // Save the workbook (lifecycle: save)
        workbook.Save("ShowFormulasDemo.xlsx");
    }
}