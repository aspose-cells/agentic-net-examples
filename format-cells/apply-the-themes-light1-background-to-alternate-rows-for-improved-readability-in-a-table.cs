using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class ApplyLight1RowStripes
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (including header row)
        worksheet.Cells["A1"].PutValue("Item");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("Cherry");
        worksheet.Cells["B4"].PutValue(15);
        worksheet.Cells["A5"].PutValue("Date");
        worksheet.Cells["B5"].PutValue(8);

        // Add a ListObject (table) covering the data range
        int tableIndex = worksheet.ListObjects.Add("A1", "B5", true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Apply the built‑in Light1 table style
        table.TableStyleType = TableStyleType.TableStyleLight1;

        // Enable row stripe formatting (alternating background)
        table.ShowTableStyleRowStripes = true;

        // Save the workbook
        workbook.Save("TableWithLight1RowStripes.xlsx", SaveFormat.Xlsx);
    }
}