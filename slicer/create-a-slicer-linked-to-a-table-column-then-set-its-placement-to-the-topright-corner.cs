using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill sample data for the table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("A");
        sheet.Cells["B4"].PutValue(30);

        // Add a ListObject (table) covering the data range
        int tableIdx = sheet.ListObjects.Add(0, 0, 3, 1, true);
        ListObject table = sheet.ListObjects[tableIdx];

        // Add a slicer linked to the first column of the table.
        // Position it initially at row 1, column 5 (near the top‑right of the sheet).
        int slicerIdx = sheet.Slicers.Add(table, table.ListColumns[0], 1, 5);
        Slicer slicer = sheet.Slicers[slicerIdx];

        // Set the slicer's placement to free‑floating so it can be positioned
        // independently of the cells (top‑right corner of the worksheet).
        slicer.Placement = PlacementType.FreeFloating;

        // Fine‑tune the exact location (optional).
        slicer.TopPixel = 0;      // align with the top edge
        slicer.LeftPixel = 500;   // shift toward the right edge (adjust as needed)

        // Save the workbook
        workbook.Save("SlicerTopRight.xlsx");
    }
}