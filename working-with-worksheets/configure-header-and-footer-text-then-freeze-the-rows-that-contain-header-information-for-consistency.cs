using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data (first row will act as the header)
        worksheet.Cells["A1"].PutValue("Report Title");
        worksheet.Cells["B1"].PutValue("Generated On");
        worksheet.Cells["A2"].PutValue("Item");
        worksheet.Cells["B2"].PutValue("Quantity");
        worksheet.Cells["A3"].PutValue("Apple");
        worksheet.Cells["B3"].PutValue(120);
        worksheet.Cells["A4"].PutValue("Orange");
        worksheet.Cells["B4"].PutValue(85);

        // ----- Configure Header -----
        // Left section
        worksheet.PageSetup.SetHeader(0, "&\"Arial,Bold\"&12Left Header");
        // Center section
        worksheet.PageSetup.SetHeader(1, "&\"Arial,Bold\"&12Center Header");
        // Right section
        worksheet.PageSetup.SetHeader(2, "&\"Arial,Bold\"&12Right Header");

        // ----- Configure Footer -----
        // Left section: page number
        worksheet.PageSetup.SetFooter(0, "Page &P of &N");
        // Center section: current date
        worksheet.PageSetup.SetFooter(1, "&D");
        // Right section: file name
        worksheet.PageSetup.SetFooter(2, "&F");

        // ----- Repeat header row on each printed page -----
        // This makes the first row repeat at the top of every printed page
        worksheet.PageSetup.PrintTitleRows = "$1:$1";

        // ----- Freeze the header row -----
        // Freeze panes at cell A2, freezing 1 row (the header) and 0 columns
        worksheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("HeaderFooterFreezeDemo.xlsx");
    }
}