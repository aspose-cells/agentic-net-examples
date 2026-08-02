// Title: C# Smart Marker Grouping for Master‑Detail Tables in Aspose.Cells
// Description: Demonstrates how to build a hierarchical order report in Excel using Aspose.Cells. The example creates a workbook, defines smart markers for order headers and line‑item rows, supplies a List<Order> with nested List<OrderDetail>, applies the WorkbookDesigner, processes the Group keyword to generate grouped rows, and saves the file as MasterDetailGrouped.xlsx.
// Keywords: Aspose.Cells | C# | .NET | smart markers | grouping | master detail | hierarchical table | WorkbookDesigner | Excel export | order report
// Common Searches: Aspose.Cells smart marker group master detail | C# create hierarchical rows with smart markers | group rows using Aspose.Cells WorkbookDesigner | generate master‑detail Excel report .NET | smart marker Group keyword example
// Developer Intent: Create an Excel workbook that lists orders with their line items as grouped, hierarchical rows using Aspose.Cells smart markers in C#.
// Use Cases: Export invoices where each order header is followed by its product lines in a single sheet. | Produce a sales summary that groups orders by date with expandable detail rows for analysis. | Generate purchase‑order listings for downstream processing with master‑detail grouping.
// AI Prompts: Show how to add a subtotal row after each order using smart markers. | Provide code to apply different styles to master rows versus detail rows. | Explain how to enable expand/collapse functionality for grouped rows in the generated Excel file.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsSmartMarkerGroupingDemo
{
    // Master class representing an order
    // Demonstrates how to build a hierarchical order report in Excel using Aspose.Cells. The example creates a workbook, defines smart markers for order headers and line‑item rows, supplies a List<Order> with nested List<OrderDetail>, applies the WorkbookDesigner, processes the Group keyword to generate grouped rows, and saves the file as MasterDetailGrouped.xlsx.
    public class Order
    {
        public int OrderID { get; set; }
        public DateTime OrderDate { get; set; }
        public List<OrderDetail> Details { get; set; }

        public Order(int id, DateTime date, List<OrderDetail> details)
        {
            OrderID = id;
            OrderDate = date;
            Details = details;
        }
    }

    // Detail class representing a line item in an order
    public class OrderDetail
    {
        public string Product { get; set; }
        public int Quantity { get; set; }
        public double UnitPrice { get; set; }

        public OrderDetail(string product, int qty, double price)
        {
            Product = product;
            Quantity = qty;
            UnitPrice = price;
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Define smart markers for master‑detail grouping
            // Header row
            cells["A1"].PutValue("Order ID");
            cells["B1"].PutValue("Order Date");
            cells["C1"].PutValue("Product");
            cells["D1"].PutValue("Quantity");
            cells["E1"].PutValue("Unit Price");

            // Master row (order level) – will be repeated for each order
            cells["A2"].PutValue("&=Orders.OrderID");
            cells["B2"].PutValue("&=Orders.OrderDate");

            // Detail rows – will be repeated for each detail under the current order
            // The special "Group" keyword tells the designer to start a new group for each master record
            cells["C3"].PutValue("&=Orders.Details.Product");
            cells["D3"].PutValue("&=Orders.Details.Quantity");
            cells["E3"].PutValue("&=Orders.Details.UnitPrice");

            // 3. Prepare master‑detail data source
            List<Order> orders = new List<Order>
            {
                new Order(
                    1001,
                    new DateTime(2023, 1, 15),
                    new List<OrderDetail>
                    {
                        new OrderDetail("Laptop", 2, 1200.00),
                        new OrderDetail("Mouse", 5, 25.50)
                    }),

                new Order(
                    1002,
                    new DateTime(2023, 2, 3),
                    new List<OrderDetail>
                    {
                        new OrderDetail("Smartphone", 3, 800.00),
                        new OrderDetail("Headphones", 4, 45.75),
                        new OrderDetail("Charger", 6, 15.00)
                    })
            };

            // 4. Set the data source and process smart markers (rule usage)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Orders", orders);
            designer.Process(); // processes all smart markers, creates hierarchical rows

            // 5. Save the result (lifecycle rule)
            workbook.Save("MasterDetailGrouped.xlsx");
        }
    }
}
