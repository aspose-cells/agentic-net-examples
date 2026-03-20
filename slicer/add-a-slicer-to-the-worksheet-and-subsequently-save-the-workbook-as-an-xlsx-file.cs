using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

class AddSlicerAndSave
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for a table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("A");
        worksheet.Cells["B4"].PutValue(30);

        // Add a ListObject (table) covering the data range
        int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Add a slicer for the first column of the table, placing it at cell D1
        int slicerIndex = worksheet.Slicers.Add(table, 0, "D1");
        Slicer slicer = worksheet.Slicers[slicerIndex];
        slicer.Caption = "Category Slicer";

        // Save the workbook as an XLSX file
        workbook.Save("SlicerDemo.xlsx", SaveFormat.Xlsx);
    }
}