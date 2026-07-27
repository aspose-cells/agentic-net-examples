using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Original nested object model
    public class Order
    {
        public int OrderId { get; set; }
        public string CustomerName { get; set; }
        public List<OrderItem> Items { get; set; }
    }

    public class OrderItem
    {
        public string ProductName { get; set; }
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
    }

    // Flattened DTO used for import
    public class FlatOrderItem
    {
        public int OrderId { get; set; }
        public string CustomerName { get; set; }
        public string ProductName { get; set; }
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
    }

    public class ImportNestedObjectsDemo
    {
        public static void Run()
        {
            // Prepare sample data with nested collections
            var orders = new List<Order>
            {
                new Order
                {
                    OrderId = 1001,
                    CustomerName = "Alice",
                    Items = new List<OrderItem>
                    {
                        new OrderItem { ProductName = "Apple", Quantity = 10, UnitPrice = 0.5m },
                        new OrderItem { ProductName = "Banana", Quantity = 5, UnitPrice = 0.3m }
                    }
                },
                new Order
                {
                    OrderId = 1002,
                    CustomerName = "Bob",
                    Items = new List<OrderItem>
                    {
                        new OrderItem { ProductName = "Orange", Quantity = 8, UnitPrice = 0.4m },
                        new OrderItem { ProductName = "Grapes", Quantity = 3, UnitPrice = 1.2m }
                    }
                }
            };

            // Flatten nested collections into a list of simple objects
            var flatItems = new List<FlatOrderItem>();
            foreach (var order in orders)
            {
                foreach (var item in order.Items)
                {
                    flatItems.Add(new FlatOrderItem
                    {
                        OrderId = order.OrderId,
                        CustomerName = order.CustomerName,
                        ProductName = item.ProductName,
                        Quantity = item.Quantity,
                        UnitPrice = item.UnitPrice
                    });
                }
            }

            // Define the order of columns to be imported
            string[] propertyNames = { "OrderId", "CustomerName", "ProductName", "Quantity", "UnitPrice" };

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Import the flattened data into the worksheet
            // Parameters:
            //   list               : flatItems
            //   propertyNames      : propertyNames (null would import all properties)
            //   isPropertyNameShown: true (adds header row)
            //   firstRow           : 0 (start at first row)
            //   firstColumn        : 0 (start at first column)
            //   rowNumber          : flatItems.Count
            //   insertRows         : true (adds rows if needed)
            //   dateFormatString   : "yyyy-MM-dd" (not used here but required)
            //   convertStringToNumber: true (convert numeric strings)
            worksheet.Cells.ImportCustomObjects(
                flatItems,
                propertyNames,
                true,
                0,
                0,
                flatItems.Count,
                true,
                "yyyy-MM-dd",
                true
            );

            // Save the workbook to a file
            workbook.Save("FlattenedOrders.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            ImportNestedObjectsDemo.Run();
        }
    }
}