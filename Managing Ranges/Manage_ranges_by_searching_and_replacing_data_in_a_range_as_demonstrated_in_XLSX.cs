using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class ManageRangeReplaceDemo
{
    static void Main()
    {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        cells["A1"].PutValue("Apple");
        cells["B1"].PutValue("Banana");
        cells["C1"].PutValue("Apple");
        cells["A2"].PutValue("Orange");
        cells["B2"].PutValue("Apple");
        cells["C2"].PutValue("Grape");
        cells["A3"].PutValue("Apple");
        cells["B3"].PutValue("Lemon");
        cells["C3"].PutValue("Apple");

        AsposeRange range = cells.CreateRange("A1", "C3");

        string oldText = "Apple";
        string newText = "Mango";

        for (int i = 0; i < range.RowCount; i++)
        {
            for (int j = 0; j < range.ColumnCount; j++)
            {
                Cell cell = range[i, j];
                if (cell.StringValue == oldText)
                {
                    cell.PutValue(newText);
                }
            }
        }

        workbook.Save("RangeReplaceDemo.xlsx");
    }
}