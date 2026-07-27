using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

class AdjustSlicerRowHeight
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a table with at least two columns
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Item1");
        sheet.Cells["B2"].PutValue("CatA");
        sheet.Cells["A3"].PutValue("Item2");
        sheet.Cells["B3"].PutValue("CatB");
        sheet.Cells["A4"].PutValue("Item3");
        sheet.Cells["B4"].PutValue("CatA");

        // Add a ListObject (table) covering the data range
        int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
        ListObject table = sheet.ListObjects[tableIndex];
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Add a slicer linked to the second column (index 1) of the table
        // Place the slicer starting at row 6, column 1 (cell B6)
        int slicerIndex = sheet.Slicers.Add(table, table.ListColumns[1], 6, 1);
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Adjust the row height of each slicer item to 30 points
        slicer.RowHeight = 30;

        // Save the workbook
        workbook.Save("AdjustedSlicerRowHeight.xlsx");
    }
}