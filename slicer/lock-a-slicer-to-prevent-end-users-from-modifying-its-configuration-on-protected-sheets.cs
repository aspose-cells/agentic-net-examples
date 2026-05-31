using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

class LockSlicerDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("A");
        worksheet.Cells["A5"].PutValue("B");

        // Create a table from the data range
        int tableIndex = worksheet.ListObjects.Add("A1", "A5", true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Add a slicer linked to the first column of the table
        int slicerIndex = worksheet.Slicers.Add(table, 0, "C1");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Lock the slicer so its configuration cannot be changed when the sheet is protected
        slicer.IsLocked = true;          // Locks the slicer shape (obsolete but functional)
        slicer.LockedPosition = true;    // Prevents moving or resizing the slicer

        // Protect the worksheet with all protection options
        worksheet.Protect(ProtectionType.All, "password123", null);

        // Save the workbook
        workbook.Save("LockedSlicerDemo.xlsx");
    }
}