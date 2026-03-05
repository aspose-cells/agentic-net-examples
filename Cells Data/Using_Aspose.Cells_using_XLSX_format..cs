using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(1.5);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(0.75);
        sheet.Cells["A4"].PutValue("Orange");
        sheet.Cells["B4"].PutValue(1.25);

        // Configure save options for XLSX (save rule)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            // Disabling cell name export improves performance and reduces file size
            ExportCellName = false
        };

        // Save the workbook to an XLSX file using the specified options (save rule)
        workbook.Save("Products.xlsx", saveOptions);

        Console.WriteLine("Workbook saved successfully as Products.xlsx");
    }
}