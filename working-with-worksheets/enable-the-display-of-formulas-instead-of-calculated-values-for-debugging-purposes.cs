using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a formula in cell A1
        worksheet.Cells["A1"].Formula = "=1+2+3";

        // Enable display of formulas instead of calculated values (debug mode)
        worksheet.ShowFormulas = true;

        // Save the workbook (lifecycle: save)
        workbook.Save("DebugFormulas.xlsx");
    }
}