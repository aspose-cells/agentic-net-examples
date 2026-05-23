using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsMasterDetailSmartMarkers
{
    // Sample data classes
    public class Order
    {
        public int OrderID { get; set; }
        public List<Item> Items { get; set; }
    }

    public class Item
    {
        public string Product { get; set; }
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ---------- Set up smart markers ----------
            // Parent marker (order header) placed above child marker range
            cells["A1"].PutValue("&=Orders.OrderID");          // Parent smart marker
            cells["B1"].PutValue("Order ID");                 // Header for readability

            // Child marker range (order items)
            // These markers will be repeated for each item of the current order
            cells["A2"].PutValue("&=Orders.Items.Product");   // Child smart marker
            cells["B2"].PutValue("&=Orders.Items.Quantity");  // Child smart marker

            // ---------- Prepare sample data ----------
            var orders = new List<Order>
            {
                new Order
                {
                    OrderID = 1001,
                    Items = new List<Item>
                    {
                        new Item { Product = "Apple",  Quantity = 5 },
                        new Item { Product = "Banana", Quantity = 3 }
                    }
                },
                new Order
                {
                    OrderID = 1002,
                    Items = new List<Item>
                    {
                        new Item { Product = "Orange", Quantity = 2 },
                        new Item { Product = "Grapes", Quantity = 4 },
                        new Item { Product = "Mango",  Quantity = 1 }
                    }
                }
            };

            // ---------- Process smart markers ----------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // Use range smart markers (LineByLine is obsolete, but kept for compatibility)
                LineByLine = false
            };
            designer.SetDataSource("Orders", orders);
            designer.Process(); // Expands rows for master‑detail data

            // ---------- Group detail rows under each order ----------
            // After processing, the sheet will look like:
            // Row 0: Header (Order ID)
            // Row 1: Order 1001
            // Row 2..3: Items of Order 1001
            // Row 4: Order 1002
            // Row 5..7: Items of Order 1002
            // We'll group the item rows for each order.

            int lastRow = cells.MaxDataRow; // Total rows after processing
            int currentRow = 1; // Start after header row

            while (currentRow <= lastRow)
            {
                // The order row contains a numeric OrderID (parent row)
                object orderCellValue = cells[currentRow, 0].Value;
                if (orderCellValue != null && int.TryParse(orderCellValue.ToString(), out _))
                {
                    int orderRow = currentRow;
                    int detailStart = orderRow + 1;
                    int detailEnd = detailStart - 1;

                    // Scan subsequent rows until we hit the next order or end of sheet
                    int scanRow = detailStart;
                    while (scanRow <= lastRow)
                    {
                        object cellVal = cells[scanRow, 0].Value;
                        // If the cell contains a numeric value, it's the next order header
                        if (cellVal != null && int.TryParse(cellVal.ToString(), out _))
                            break;

                        // Otherwise it's a detail row
                        detailEnd = scanRow;
                        scanRow++;
                    }

                    // If there are detail rows, group them
                    if (detailEnd >= detailStart)
                    {
                        // Group rows (detailStart to detailEnd) and hide them initially
                        cells.GroupRows(detailStart, detailEnd, true);
                        // Optionally set outline to show summary row below details
                        sheet.Outline.SummaryRowBelow = true;
                    }

                    // Move to the next order row
                    currentRow = scanRow;
                }
                else
                {
                    // Safety fallback to avoid infinite loop
                    currentRow++;
                }
            }

            // ---------- Save the result ----------
            workbook.Save("MasterDetailGrouped.xlsx");
        }
    }
}