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

        // Add sample data for the table
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["A4"].PutValue("Cherry");

        // Create a ListObject (table) covering the data range (A1:A4)
        int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 0, true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Apply a compact style (optional for dashboard appearance)
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Hide the header row to achieve a clean dashboard view
        table.ShowHeaderRow = false;

        // Save the workbook
        workbook.Save("DashboardTable.xlsx");
    }
}