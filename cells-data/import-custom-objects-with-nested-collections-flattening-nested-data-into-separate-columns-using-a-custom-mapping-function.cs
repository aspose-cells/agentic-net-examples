using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportNestedObjectsDemo
{
    // Domain classes with nested collections
    public class Order
    {
        public int OrderId { get; set; }
        public Customer Customer { get; set; }
        public List<Item> Items { get; set; }
    }

    public class Customer
    {
        public string Name { get; set; }
        public string Email { get; set; }
    }

    public class Item
    {
        public string Product { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
    }

    // Flat DTO used for import
    public class FlatOrderItem
    {
        public int OrderId { get; set; }
        public string CustomerName { get; set; }
        public string CustomerEmail { get; set; }
        public string Product { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
    }

    public static class Program
    {
        // Custom mapping function: flattens nested collections into a list of flat objects
        private static List<FlatOrderItem> FlattenOrders(IEnumerable<Order> orders)
        {
            var flatList = new List<FlatOrderItem>();

            foreach (var order in orders)
            {
                if (order.Items == null) continue;

                foreach (var item in order.Items)
                {
                    flatList.Add(new FlatOrderItem
                    {
                        OrderId = order.OrderId,
                        CustomerName = order.Customer?.Name,
                        CustomerEmail = order.Customer?.Email,
                        Product = item.Product,
                        Quantity = item.Quantity,
                        Price = item.Price
                    });
                }
            }

            return flatList;
        }

        public static void Main()
        {
            // Sample data with nested collections
            var orders = new List<Order>
            {
                new Order
                {
                    OrderId = 1001,
                    Customer = new Customer { Name = "Alice", Email = "alice@example.com" },
                    Items = new List<Item>
                    {
                        new Item { Product = "Laptop", Quantity = 1, Price = 1200.00m },
                        new Item { Product = "Mouse", Quantity = 2, Price = 25.50m }
                    }
                },
                new Order
                {
                    OrderId = 1002,
                    Customer = new Customer { Name = "Bob", Email = "bob@example.com" },
                    Items = new List<Item>
                    {
                        new Item { Product = "Keyboard", Quantity = 1, Price = 75.00m }
                    }
                }
            };

            // Flatten the nested structure
            List<FlatOrderItem> flatData = FlattenOrders(orders);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the order of columns to be imported
            string[] propertyNames = new[]
            {
                "OrderId",
                "CustomerName",
                "CustomerEmail",
                "Product",
                "Quantity",
                "Price"
            };

            // Import the flat data into the worksheet
            // Parameters:
            // list, propertyNames, isPropertyNameShown, firstRow, firstColumn,
            // rowNumber, insertRows, dateFormatString, convertStringToNumber
            worksheet.Cells.ImportCustomObjects(
                flatData,
                propertyNames,
                true,          // show column headers
                0,             // start at first row (A1)
                0,             // start at first column
                flatData.Count,
                true,          // insert rows if needed
                "yyyy-MM-dd",  // date format (not used here but required)
                true           // try to convert strings to numbers
            );

            // Save the workbook
            workbook.Save("FlattenedOrders.xlsx");
        }
    }
}