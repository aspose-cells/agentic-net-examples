using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate data for the table (including header row)
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("Cherry");
        worksheet.Cells["B4"].PutValue(30);

        // Add a ListObject (structured table) covering the range A1:B4, with headers
        int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.TableStyleType = TableStyleType.TableStyleMedium2; // optional styling

        // Freeze the header row so it stays visible while scrolling
        // Freeze at cell A2 (row index 2), freezing 1 row (the header) and 0 columns
        worksheet.FreezePanes(2, 0, 1, 0);

        // Save the workbook
        workbook.Save("TableWithFrozenHeader.xlsx");
    }
}