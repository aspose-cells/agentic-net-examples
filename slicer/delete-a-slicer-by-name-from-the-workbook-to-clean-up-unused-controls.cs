using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

class DeleteSlicerByName
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some data and create a table (required for slicer source)
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("A");

        int tableIndex = worksheet.ListObjects.Add("A1", "A4", true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Add a slicer linked to the first column of the table
        int slicerIndex = worksheet.Slicers.Add(table, 0, "C1");
        Slicer slicer = worksheet.Slicers[slicerIndex];
        slicer.Name = "MySlicer"; // Assign a custom name to the slicer

        // Retrieve the slicer by its name and remove it from the collection
        Slicer slicerToRemove = worksheet.Slicers["MySlicer"];
        worksheet.Slicers.Remove(slicerToRemove);

        // Save the workbook
        workbook.Save("DeletedSlicer.xlsx");
    }
}