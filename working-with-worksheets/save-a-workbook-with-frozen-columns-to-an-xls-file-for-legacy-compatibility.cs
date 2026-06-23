using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill some sample data
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Score");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["C2"].PutValue(85);
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");
        sheet.Cells["C3"].PutValue(92);

        // Freeze the first two columns (A and B) by freezing at cell C1
        // 0 frozen rows, 2 frozen columns
        sheet.FreezePanes("C1", 0, 2);

        // Create XLS save options for legacy Excel 97-2003 format
        XlsSaveOptions saveOptions = new XlsSaveOptions();

        // Save the workbook as an XLS file with the frozen columns
        workbook.Save("FrozenColumns.xls", saveOptions);
    }
}