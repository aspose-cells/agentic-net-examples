using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace ListObjectBenefitsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["D1"].PutValue("Quantity");

            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue("Electronics");
            sheet.Cells["C2"].PutValue(1200);
            sheet.Cells["D2"].PutValue(5);

            sheet.Cells["A3"].PutValue("Desk");
            sheet.Cells["B3"].PutValue("Furniture");
            sheet.Cells["C3"].PutValue(300);
            sheet.Cells["D3"].PutValue(12);

            sheet.Cells["A4"].PutValue("Pen");
            sheet.Cells["B4"].PutValue("Stationery");
            sheet.Cells["C4"].PutValue(2);
            sheet.Cells["D4"].PutValue(150);

            // Add a ListObject (Excel table) covering A1:D4
            int tableIndex = sheet.ListObjects.Add(0, 0, 3, 3, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "InventoryTable";
            table.HasAutoFilter = true;
            table.ShowTotals = true;
            table.ListColumns[2].TotalsCalculation = TotalsCalculation.Sum;
            table.ListColumns[3].TotalsCalculation = TotalsCalculation.Sum;
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Add more rows
            sheet.Cells["A5"].PutValue("Notebook");
            sheet.Cells["B5"].PutValue("Stationery");
            sheet.Cells["C5"].PutValue(5);
            sheet.Cells["D5"].PutValue(80);

            // Resize the table to include the new row
            table.Resize(0, 0, 4, 3, true);

            // Retrieve the table by name
            ListObject retrievedTable = sheet.ListObjects["InventoryTable"];
            retrievedTable.ShowHeaderRow = true;

            // Access the data range
            AsposeRange dataRange = retrievedTable.DataRange;
            Console.WriteLine($"Data range address: {dataRange.RefersTo}");
            Console.WriteLine($"Rows: {dataRange.RowCount}, Columns: {dataRange.ColumnCount}");

            // Apply number format to the Price column
            for (int r = 0; r < dataRange.RowCount; r++)
            {
                Cell cell = dataRange[r, 2];
                Style style = cell.GetStyle();
                style.Number = 2; // Two decimal places
                cell.SetStyle(style);
            }

            // Save the workbook
            workbook.Save("ListObjectBenefitsDemo.xlsx");
            Console.WriteLine("Workbook with ListObject created successfully.");
        }
    }
}