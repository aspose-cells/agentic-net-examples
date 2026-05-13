using System;
using System.Text;
using Aspose.Cells;

class ExportTabDelimited
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["A2"].PutValue("Laptop");
        sheet.Cells["B2"].PutValue(999.99);
        sheet.Cells["A3"].PutValue("Phone");
        sheet.Cells["B3"].PutValue(599.99);
        sheet.Cells["A4"].PutValue("Tablet");
        sheet.Cells["B4"].PutValue(399.99);

        // Configure text save options for tab-delimited output
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.Separator = '\t'; // Tab character as delimiter
        saveOptions.Encoding = Encoding.UTF8;

        // Save the workbook as a tab-delimited TXT file
        workbook.Save("Products.txt", saveOptions);
    }
}