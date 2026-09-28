// Title: Generate hierarchical master‑detail Excel tables using Aspose.Cells smart marker grouping in C#
// AI Prompts: Write C# code that defines Order and OrderDetail classes, creates a workbook template with smart markers, binds a List<Order> to WorkbookDesigner, and processes the markers to produce a grouped master‑detail sheet. | Show how to add a subtotal row for each order's detail section using smart marker syntax and calculate totals automatically. | Demonstrate how to export the processed workbook to a .xlsx file and open it for verification.
// Common Searches: aspocells c# smart markers hierarchical master detail grouping | how to use smart marker syntax for nested collections in Excel with Aspose.Cells | C# generate master‑detail Excel report with row grouping using Aspose.Cells | smart marker template for orders and order details Aspose.Cells .NET | group master rows and detail rows automatically Aspose.Cells smart markers
// Tags: smart marker grouping master‑detail | C# Aspose.Cells hierarchical export | bind nested collection WorkbookDesigner | Excel row grouping with smart markers | generate master‑detail .xlsx using Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsSmartMarkerGroupingDemo
{
    // Master class
    // The example creates a workbook, inserts smart markers for Order master rows and OrderDetail rows, builds a List<Order> with nested OrderDetail collections, sets this list as the "Orders" data source for a WorkbookDesigner, processes the markers to automatically expand and group master‑detail rows, and saves the result as MasterDetailGrouped.xlsx.
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

    // Detail class
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
            // 1. Create a new workbook (template) and add smart markers.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row for master data
            cells["A1"].PutValue("Order ID");
            cells["B1"].PutValue("Order Date");

            // Master smart markers (first row of data)
            cells["A2"].PutValue("&=$Orders.OrderID");
            cells["B2"].PutValue("&=$Orders.OrderDate");

            // Header row for detail data (placed directly below master row)
            cells["A3"].PutValue("Product");
            cells["B3"].PutValue("Quantity");
            cells["C3"].PutValue("Unit Price");

            // Detail smart markers (repeat for each order)
            cells["A4"].PutValue("&=$Orders.Details.Product");
            cells["B4"].PutValue("&=$Orders.Details.Quantity");
            cells["C4"].PutValue("&=$Orders.Details.UnitPrice");

            // 2. Prepare hierarchical data (master‑detail).
            List<Order> orders = new List<Order>
            {
                new Order(
                    1001,
                    new DateTime(2023, 1, 15),
                    new List<OrderDetail>
                    {
                        new OrderDetail("Apple", 10, 1.5),
                        new OrderDetail("Banana", 5, 0.9)
                    })
                ,
                new Order(
                    1002,
                    new DateTime(2023, 2, 3),
                    new List<OrderDetail>
                    {
                        new OrderDetail("Orange", 8, 1.2),
                        new OrderDetail("Grapes", 12, 2.0),
                        new OrderDetail("Mango", 3, 3.5)
                    })
            };

            // 3. Bind the data source to the workbook designer using the smart marker name.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Orders", orders);

            // 4. Process the smart markers – Aspose.Cells will automatically expand the master rows,
            //    insert detail rows for each master, and maintain the hierarchical grouping.
            designer.Process();

            // 5. Save the result.
            workbook.Save("MasterDetailGrouped.xlsx");
        }
    }
}
