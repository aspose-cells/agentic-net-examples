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

        // Add sample data
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("Item1");
        worksheet.Cells["B2"].PutValue(2);
        worksheet.Cells["A3"].PutValue("Item2");
        worksheet.Cells["B3"].PutValue(3);

        // Set the print area to the range A1:B3
        worksheet.PageSetup.PrintArea = "A1:B3";

        // Save the workbook as an Excel file
        workbook.Save("PrintAreaDemo.xlsx");

        // Save the workbook as HTML, exporting only the defined print area
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportPrintAreaOnly = true
        };
        workbook.Save("PrintAreaDemo.html", htmlOptions);
    }
}