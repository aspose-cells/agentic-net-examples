using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Food");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("Drink");
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["A4"].PutValue("Food");
        sheet.Cells["B4"].PutValue(200);
        sheet.Cells["A5"].PutValue("Snack");
        sheet.Cells["B5"].PutValue(80);

        // Add a ListObject (table) that covers the data range A1:B5
        int tableIndex = sheet.ListObjects.Add("A1", "B5", true);
        ListObject table = sheet.ListObjects[tableIndex];
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Add a slicer linked to the "Category" column (first column of the table)
        // Place the slicer with its upper‑left corner at cell A7 (row index 6, column index 0)
        SlicerCollection slicers = sheet.Slicers;
        int slicerIndex = slicers.Add(table, table.ListColumns[0], 6, 0);
        Slicer slicer = slicers[slicerIndex];

        // Optional: customize slicer appearance
        slicer.Caption = "Category Filter";
        slicer.NumberOfColumns = 1;
        slicer.WidthPixel = 150;
        slicer.HeightPixel = 120;

        // Save the workbook
        workbook.Save("SlicerTableDemo.xlsx");
    }
}