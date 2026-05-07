using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsMultipleRangesDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // First data section: Sales data (A1:C5)
            // ------------------------------------------------------------
            // Create a range using address string
            AsposeRange salesRange = cells.CreateRange("A1", "C5");
            // Fill the range with sample data
            salesRange[0, 0].PutValue("Product");
            salesRange[0, 1].PutValue("Region");
            salesRange[0, 2].PutValue("Sales");
            salesRange[1, 0].PutValue("Widget");
            salesRange[1, 1].PutValue("North");
            salesRange[1, 2].PutValue(1200);
            salesRange[2, 0].PutValue("Gadget");
            salesRange[2, 1].PutValue("South");
            salesRange[2, 2].PutValue(850);
            // Add the range to the Cells' internal collection (so it expands with inserts)
            cells.AddRange(salesRange);

            // ------------------------------------------------------------
            // Second data section: Inventory data (E1:G4)
            // ------------------------------------------------------------
            // Create a range using integer coordinates (startRow, startColumn, totalRows, totalColumns)
            // E column = index 4, G column = index 6
            AsposeRange inventoryRange = cells.CreateRange(0, 4, 4, 3);
            // Fill the range with sample data
            inventoryRange[0, 0].PutValue("Item");
            inventoryRange[0, 1].PutValue("Qty");
            inventoryRange[0, 2].PutValue("Warehouse");
            inventoryRange[1, 0].PutValue("Screws");
            inventoryRange[1, 1].PutValue(500);
            inventoryRange[1, 2].PutValue("WH1");
            inventoryRange[2, 0].PutValue("Nuts");
            inventoryRange[2, 1].PutValue(300);
            inventoryRange[2, 2].PutValue("WH2");
            // Add this second range to the Cells collection as well
            cells.AddRange(inventoryRange);

            // ------------------------------------------------------------
            // Demonstrate that both ranges are tracked independently.
            // Insert a row at index 2 (third row). Both ranges will expand automatically.
            // ------------------------------------------------------------
            cells.InsertRows(2, 1);

            // After insertion, update a value in the expanded sales range to show it shifted.
            // The original row 2 (index 1) moved to index 2.
            salesRange[2, 2].PutValue(1300); // Update Sales value for the moved row

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("MultipleRangesDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}