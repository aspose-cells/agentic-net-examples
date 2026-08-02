using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with a "Region" column
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Region");
        worksheet.Cells["A2"].PutValue("Laptop");
        worksheet.Cells["B2"].PutValue("East");
        worksheet.Cells["A3"].PutValue("Phone");
        worksheet.Cells["B3"].PutValue("West");
        worksheet.Cells["A4"].PutValue("Monitor");
        worksheet.Cells["B4"].PutValue("East");
        worksheet.Cells["A5"].PutValue("Keyboard");
        worksheet.Cells["B5"].PutValue("South");

        // Create a ListObject (table) that covers the data range A1:B5
        int listObjectIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
        ListObject listObject = worksheet.ListObjects[listObjectIndex];

        // Enable auto‑filter for the table
        listObject.HasAutoFilter = true;

        // Apply a filter on the "Region" column (index 1) to show only rows where Region = "East"
        if (listObject.AutoFilter != null)
        {
            listObject.AutoFilter.Filter(1, "East");
            listObject.AutoFilter.Refresh();
        }

        // Save the workbook
        workbook.Save("FilteredListObject.xlsx");
    }
}