using System;
using Aspose.Cells;

class FreezeDemo
{
    static void Main()
    {
        // Workbook is created inside a using block so Dispose() is called automatically
        using (Workbook workbook = new Workbook())
        {
            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B1"].PutValue("Data");
            sheet.Cells["B2"].PutValue(456);

            // Freeze panes at cell C3 (2 rows and 2 columns are frozen)
            sheet.FreezePanes("C3", 2, 2);

            // Save the workbook to disk
            workbook.Save("FrozenDemo.xlsx");
        } // workbook.Dispose() is invoked here, guaranteeing proper resource cleanup
    }
}