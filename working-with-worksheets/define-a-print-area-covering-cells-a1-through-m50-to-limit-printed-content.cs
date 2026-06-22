using System;
using Aspose.Cells;

class SetPrintAreaDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the print area to cover cells A1 through M50
        worksheet.PageSetup.PrintArea = "A1:M50";

        // Save the workbook (optional, demonstrates that the print area is applied)
        workbook.Save("PrintAreaDemo.xlsx");
    }
}