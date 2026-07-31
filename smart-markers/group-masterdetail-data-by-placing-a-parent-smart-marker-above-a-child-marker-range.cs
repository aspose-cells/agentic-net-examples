// Title: C# Aspose.Cells – Master‑Detail Smart Markers with Collapsible Detail Rows
// Description: This example shows how to build an Excel workbook in C# using Aspose.Cells smart markers where a parent marker (OrderId) sits above a child marker block (Product, Qty). The marker range is named, fed with master (Orders) and detail (Items) collections, processed with WorkbookDesigner, and then the generated detail rows are programmatically grouped and collapsed under each master row before saving the file.
// Keywords: Aspose.Cells | C# | smart markers | master detail | row grouping | collapse rows | WorkbookDesigner | named range | Excel export | hierarchical data
// Common Searches: Aspose.Cells master detail smart markers C# example | group child rows under a parent marker in Excel using Aspose | process a specific smart‑marker range with WorkbookDesigner | collapse detail rows after smart marker processing | create expandable sections in Excel with Aspose.Cells
// Developer Intent: Create an Excel file where each order header appears as a master row and its line items are hidden under a collapsible group, using smart markers and programmatic row grouping.
// Use Cases: Generating invoices where each order header can be expanded to reveal line items. | Building sales dashboards that let users drill down from region totals to individual product sales. | Exporting nested .NET objects (e.g., orders and items) to a readable, collapsible Excel report.
// AI Prompts: Modify the code to determine rowsPerOrder dynamically based on the actual item count for each order. | Show how to rename the child data source (e.g., to OrderItems) while keeping the same smart‑marker template. | Demonstrate applying a custom style to master rows after the smart‑marker processing completes.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsMasterDetailSmartMarkers
{
    // Simple data models for master‑detail relationship
    // This example shows how to build an Excel workbook in C# using Aspose.Cells smart markers where a parent marker (OrderId) sits above a child marker block (Product, Qty). The marker range is named, fed with master (Orders) and detail (Items) collections, processed with WorkbookDesigner, and then the generated detail rows are programmatically grouped and collapsed under each master row before saving the file.
    public class Order
    {
        public string OrderId { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Product { get; set; } = string.Empty;
        public int Qty { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Build a template with smart markers
                // Row 0 – parent smart marker (master)
                cells["A1"].PutValue("&=Orders.OrderId");
                // Row 1‑2 – child smart markers (detail)
                cells["A2"].PutValue("&=Items.Product");
                cells["B2"].PutValue("&=Items.Qty");
                cells["A3"].PutValue("&=Items.Product");
                cells["B3"].PutValue("&=Items.Qty");

                // Define the range that contains the smart markers and give it the required name
                // This tells the designer to treat the whole block as a smart‑marker range
                Aspose.Cells.Range smartRange = cells.CreateRange("A1:B3");
                smartRange.Name = "_CellsSmartMarkers";

                // 3. Prepare sample master‑detail data
                var orders = new List<Order>
                {
                    new Order
                    {
                        OrderId = "ORD001",
                        Items = new List<Item>
                        {
                            new Item { Product = "Apple",  Qty = 10 },
                            new Item { Product = "Banana", Qty = 20 }
                        }
                    },
                    new Order
                    {
                        OrderId = "ORD002",
                        Items = new List<Item>
                        {
                            new Item { Product = "Orange", Qty = 15 },
                            new Item { Product = "Grape",  Qty = 25 }
                        }
                    }
                };

                // 4. Set up the WorkbookDesigner
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                // Set data sources for master and detail
                designer.SetDataSource("Orders", orders);
                // The detail source name must match the child marker prefix (Items)
                // Using the first order's items as a placeholder; the designer will repeat the detail rows per master row
                designer.SetDataSource("Items", orders[0].Items);

                // 5. Process the smart markers – only the defined range is processed
                designer.Process(smartRange, true);

                // 6. After processing, group detail rows under each master row
                // The processed rows start at index 0 (Excel row 1)
                // Each order expands to 1 master row + 2 detail rows (as defined in the template)
                int rowsPerOrder = 3; // 1 master + 2 detail
                for (int orderIndex = 0; orderIndex < orders.Count; orderIndex++)
                {
                    int masterRow = orderIndex * rowsPerOrder;               // zero‑based index of master row
                    int firstDetailRow = masterRow + 1;                      // first detail row
                    int lastDetailRow = masterRow + rowsPerOrder - 1;        // last detail row
                    // Group the detail rows and hide them (collapsed view)
                    cells.GroupRows(firstDetailRow, lastDetailRow, true);
                }

                // 7. Save the resulting workbook
                workbook.Save("MasterDetailSmartMarkers.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
