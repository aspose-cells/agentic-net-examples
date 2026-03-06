using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("A");
        sheet.Cells["B4"].PutValue(30);

        // Add a table that includes the data (header row = true)
        int tableIndex = sheet.ListObjects.Add(0, 0, 3, 1, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Add a slicer for the first column of the table at position (row 1, column 3)
        sheet.Slicers.Add(table, table.ListColumns[0], 1, 3);

        // Save the workbook as XLSX
        workbook.Save("SlicerDemo.xlsx", SaveFormat.Xlsx);
    }
}