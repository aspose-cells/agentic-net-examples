using System;
using Aspose.Cells;

class PrettyPrintJsonDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill the worksheet with sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Jane");
        sheet.Cells["B3"].PutValue(25);

        // Create JsonSaveOptions (pretty‑printing is applied by default)
        JsonSaveOptions saveOptions = new JsonSaveOptions();

        // Save the workbook as a JSON file with the specified options
        workbook.Save("prettyPrinted.json", saveOptions);
    }
}