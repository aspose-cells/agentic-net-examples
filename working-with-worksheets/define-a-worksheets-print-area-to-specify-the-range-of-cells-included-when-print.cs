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

        // Populate some sample data (optional)
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("Item1");
        worksheet.Cells["B2"].PutValue(2);
        worksheet.Cells["A3"].PutValue("Item2");
        worksheet.Cells["B3"].PutValue(3);

        // Define the print area for the worksheet
        worksheet.PageSetup.PrintArea = "A1:B3";

        // Save the workbook (any supported format)
        workbook.Save("PrintAreaDemo.xlsx");
    }
}