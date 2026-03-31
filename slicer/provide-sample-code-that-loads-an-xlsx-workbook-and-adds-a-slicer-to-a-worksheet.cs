using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("Input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Create sample data if the worksheet is empty
        if (sheet.Cells["A1"].Value == null)
        {
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(30);
        }

        // Add a ListObject (table) that uses the data range A1:B4
        int tableIndex = sheet.ListObjects.Add(0, 0, 3, 1, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Add a slicer for the first column of the table and place it at cell E5
        int slicerIndex = sheet.Slicers.Add(table, 0, "E5");
        Slicer slicer = sheet.Slicers[slicerIndex];
        slicer.Caption = "Category Slicer";

        // Save the workbook with the new slicer
        workbook.Save("Output.xlsx");
    }
}