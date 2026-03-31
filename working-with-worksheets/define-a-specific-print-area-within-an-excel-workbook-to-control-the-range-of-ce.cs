using System;
using Aspose.Cells;

class DefinePrintAreaDemo
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (including cells outside the intended print area)
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Item1");
        sheet.Cells["B2"].PutValue(2);
        sheet.Cells["A3"].PutValue("Item2");
        sheet.Cells["B3"].PutValue(3);
        sheet.Cells["C1"].PutValue("Extra");
        sheet.Cells["C2"].PutValue(99);

        // Define the print area to include only cells A1:B3
        sheet.PageSetup.PrintArea = "A1:B3";

        // Save the workbook; when printed, only the defined area will be considered
        workbook.Save("PrintAreaDemo.xlsx");
    }
}