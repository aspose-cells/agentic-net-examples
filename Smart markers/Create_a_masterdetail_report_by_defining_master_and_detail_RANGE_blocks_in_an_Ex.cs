using System;
using System.Collections.Generic;
using Aspose.Cells;

public class Order
{
    public int OrderID { get; set; }
    public string Customer { get; set; }
    public List<OrderItem> Items { get; set; }
}

public class OrderItem
{
    public string Product { get; set; }
    public int Quantity { get; set; }
    public double UnitPrice { get; set; }
}

class MasterDetailReport
{
    static void Main()
    {
        // Prepare sample master‑detail data
        List<Order> orders = new List<Order>
        {
            new Order
            {
                OrderID = 1001,
                Customer = "Alice",
                Items = new List<OrderItem>
                {
                    new OrderItem { Product = "Laptop", Quantity = 1, UnitPrice = 1200.0 },
                    new OrderItem { Product = "Mouse", Quantity = 2, UnitPrice = 25.5 }
                }
            },
            new Order
            {
                OrderID = 1002,
                Customer = "Bob",
                Items = new List<OrderItem>
                {
                    new OrderItem { Product = "Tablet", Quantity = 3, UnitPrice = 300.0 },
                    new OrderItem { Product = "Keyboard", Quantity = 1, UnitPrice = 45.0 }
                }
            }
        };

        // Create a new workbook and set up smart markers for master‑detail report
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Report";

        // Header row
        sheet.Cells["A1"].PutValue("Order ID");
        sheet.Cells["B1"].PutValue("Customer");
        sheet.Cells["C1"].PutValue("Product");
        sheet.Cells["D1"].PutValue("Quantity");
        sheet.Cells["E1"].PutValue("Unit Price");

        // Master markers
        sheet.Cells["A2"].PutValue("&=Order.OrderID");
        sheet.Cells["B2"].PutValue("&=Order.Customer");

        // Detail markers (will repeat for each OrderItem)
        sheet.Cells["C3"].PutValue("&=Order.Items.Product");
        sheet.Cells["D3"].PutValue("&=Order.Items.Quantity");
        sheet.Cells["E3"].PutValue("&=Order.Items.UnitPrice");

        // Initialize WorkbookDesigner and set data source
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource("Order", orders);
        designer.Process();

        // Save the populated report
        string outputPath = "MasterDetailReport.xlsx";
        workbook.Save(outputPath);
    }
}