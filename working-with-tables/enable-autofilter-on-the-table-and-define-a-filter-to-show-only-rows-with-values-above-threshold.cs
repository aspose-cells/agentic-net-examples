using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (header + 4 rows)
        worksheet.Cells["A1"].PutValue("Item");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(50);
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B3"].PutValue(30);
        worksheet.Cells["A4"].PutValue("Cherry");
        worksheet.Cells["B4"].PutValue(70);
        worksheet.Cells["A5"].PutValue("Date");
        worksheet.Cells["B5"].PutValue(20);

        // Create a ListObject (table) that covers the data range A1:B5
        int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
        ListObject listObject = worksheet.ListObjects[tableIndex];

        // Enable auto‑filter for the table
        listObject.HasAutoFilter = true;

        // Define a threshold and filter the "Quantity" column (index 1) to show rows > threshold
        int threshold = 40;
        listObject.AutoFilter.Custom(1, FilterOperatorType.GreaterThan, threshold);
        listObject.AutoFilter.Refresh();

        // Save the workbook
        workbook.Save("FilteredTable.xlsx");
    }
}