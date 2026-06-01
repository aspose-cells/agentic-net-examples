using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

class RemoveTableSlicers
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("A");
        sheet.Cells["B4"].PutValue(30);

        // Add a table (ListObject) that covers the data range
        int tableIdx = sheet.ListObjects.Add(0, 0, 3, 1, true);
        ListObject table = sheet.ListObjects[tableIdx];

        // Add a couple of slicers that are linked to the table
        SlicerCollection slicers = sheet.Slicers;
        slicers.Add(table, table.ListColumns[0], "D1"); // slicer for first column
        slicers.Add(table, table.ListColumns[1], "D5"); // slicer for second column

        // ------------------------------------------------------------
        // Remove all slicers that belong to the specified table.
        // In this example we know that the worksheet contains only
        // slicers for this table, so we can clear the entire collection.
        // If other slicers existed, you could iterate the collection and
        // remove only those whose SlicerCache is linked to the table.
        // ------------------------------------------------------------
        slicers.Clear();

        // Save the modified workbook
        workbook.Save("WorkbookWithoutTableSlicers.xlsx");
    }
}