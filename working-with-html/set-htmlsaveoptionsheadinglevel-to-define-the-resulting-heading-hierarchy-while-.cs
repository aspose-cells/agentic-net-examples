using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apples");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Oranges");
        sheet.Cells["B3"].PutValue(20);

        // Initialize HTML save options
        HtmlSaveOptions options = new HtmlSaveOptions();

        // Save the workbook as HTML (lifecycle: save)
        workbook.Save("WorkbookWithHeadings.html", options);
    }
}