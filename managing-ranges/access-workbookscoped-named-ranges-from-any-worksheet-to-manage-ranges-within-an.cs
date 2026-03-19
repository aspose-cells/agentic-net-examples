using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare worksheets and data
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0]; // default first sheet
            sheet1.Name = "Sales";

            Worksheet sheet2 = workbook.Worksheets.Add("Inventory");

            // Fill some sample data in both sheets
            sheet1.Cells["A1"].PutValue("Product");
            sheet1.Cells["B1"].PutValue("Qty");
            sheet1.Cells["A2"].PutValue("Apple");
            sheet1.Cells["B2"].PutValue(120);
            sheet1.Cells["A3"].PutValue("Banana");
            sheet1.Cells["B3"].PutValue(85);

            sheet2.Cells["A1"].PutValue("Item");
            sheet2.Cells["B1"].PutValue("Stock");
            sheet2.Cells["A2"].PutValue("Screw");
            sheet2.Cells["B2"].PutValue(500);
            sheet2.Cells["A3"].PutValue("Nail");
            sheet2.Cells["B3"].PutValue(750);

            // -------------------------------------------------
            // 2. Create workbook‑scoped named ranges
            // -------------------------------------------------
            // Named range "SalesData" refers to A1:B3 on the Sales sheet
            AsposeRange salesRange = sheet1.Cells.CreateRange("A1", "B3");
            salesRange.Name = "SalesData";

            // Named range "InventoryData" refers to A1:B3 on the Inventory sheet
            AsposeRange inventoryRange = sheet2.Cells.CreateRange("A1", "B3");
            inventoryRange.Name = "InventoryData";

            // -------------------------------------------------
            // 3. Access named ranges from any worksheet
            // -------------------------------------------------
            // Using WorksheetCollection.GetRangeByName (global scope)
            AsposeRange retrievedSales = workbook.Worksheets.GetRangeByName("SalesData");
            AsposeRange retrievedInventory = workbook.Worksheets.GetRangeByName("InventoryData");

            // Verify that the ranges were found
            if (retrievedSales != null)
            {
                Console.WriteLine($"SalesData address: {retrievedSales.Address}");
                // Example: change the quantity of the first product
                retrievedSales[1, 1].PutValue(200); // B2 cell
            }

            if (retrievedInventory != null)
            {
                Console.WriteLine($"InventoryData address: {retrievedInventory.Address}");
                // Add a new item programmatically using the worksheet cells (outside the original range)
                sheet2.Cells["A4"].PutValue("Bolt");
                sheet2.Cells["B4"].PutValue(300);
            }

            // -------------------------------------------------
            // 4. Alternative access via the Names collection
            // -------------------------------------------------
            Name salesName = workbook.Worksheets.Names["SalesData"];
            if (salesName != null)
            {
                // GetRange returns the actual Range object
                AsposeRange rangeViaName = salesName.GetRange();
                Console.WriteLine($"Accessed via Names collection: {rangeViaName.Address}");
                // Example: read a cell value
                Console.WriteLine($"First product: {rangeViaName[1, 0].StringValue}");
            }

            // -------------------------------------------------
            // 5. Save the workbook
            // -------------------------------------------------
            workbook.Save("WorkbookScopedNamedRanges.xlsx");
            Console.WriteLine("Workbook saved as WorkbookScopedNamedRanges.xlsx");
        }
    }
}