using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace HideTableHeaderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (including a header row)
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("Cherry");
            worksheet.Cells["B4"].PutValue(30);

            // Add a ListObject (table) covering the data range (including header)
            int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Apply a style (optional)
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Hide the header row while keeping data rows visible
            table.ShowHeaderRow = false;

            // Save the workbook to a file
            workbook.Save("HideTableHeaderDemo.xlsx");
        }
    }
}