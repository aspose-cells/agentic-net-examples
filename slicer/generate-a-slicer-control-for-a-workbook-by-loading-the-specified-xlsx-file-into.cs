using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        if (worksheet.ListObjects.Count == 0)
        {
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);

            worksheet.ListObjects.Add(0, 0, 2, 1, true);
        }

        ListObject table = worksheet.ListObjects[0];

        int slicerIndex = worksheet.Slicers.Add(table, 0, "D1");
        Slicer slicer = worksheet.Slicers[slicerIndex];
        slicer.Caption = "Category Slicer";

        workbook.Save("output.xlsx");
    }
}