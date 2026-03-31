using System;
using System.Collections.Generic;
using Aspose.Cells;

public class Program
{
    public static void Main(string[] args)
    {
        MasterDetailSmartMarkers.Run();
    }
}

public class MasterDetailSmartMarkers
{
    public static void Run()
    {
        // Create a new workbook and add smart markers for master‑detail data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Header for master data
        sheet.Cells["A1"].PutValue("Order ID");
        sheet.Cells["B1"].PutValue("Customer Name");

        // Master smart markers
        sheet.Cells["A2"].PutValue("&=Orders.OrderID");
        sheet.Cells["B2"].PutValue("&=Orders.CustomerName");

        // Header for detail data (starts after a blank row)
        sheet.Cells["A4"].PutValue("Product");
        sheet.Cells["B4"].PutValue("Quantity");

        // Detail smart markers (referencing the Items collection of each Order)
        sheet.Cells["A5"].PutValue("&=Orders.Items.Product");
        sheet.Cells["B5"].PutValue("&=Orders.Items.Quantity");

        // Prepare master‑detail data
        List<Order> orders = new List<Order>
        {
            new Order
            {
                OrderID = 1001,
                CustomerName = "John Doe",
                Items = new List<OrderItem>
                {
                    new OrderItem { Product = "Laptop", Quantity = 1 },
                    new OrderItem { Product = "Mouse", Quantity = 2 }
                }
            },
            new Order
            {
                OrderID = 1002,
                CustomerName = "Jane Smith",
                Items = new List<OrderItem>
                {
                    new OrderItem { Product = "Keyboard", Quantity = 1 },
                    new OrderItem { Product = "Monitor", Quantity = 2 },
                    new OrderItem { Product = "USB Cable", Quantity = 5 }
                }
            }
        };

        // Initialize the designer with the workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Set the master data source (collection of orders)
        designer.SetDataSource("Orders", orders);

        // Process the smart markers and populate the workbook
        designer.Process();

        // Save the populated workbook
        workbook.Save("MasterDetailResult.xlsx");
    }

    // Master record class
    public class Order
    {
        public int OrderID { get; set; }
        public string CustomerName { get; set; } = null!;
        public List<OrderItem> Items { get; set; } = null!;
    }

    // Detail record class
    public class OrderItem
    {
        public string Product { get; set; } = null!;
        public int Quantity { get; set; }
    }
}