// Title: C# – Flatten Nested Order Objects and Export to Excel with Aspose.Cells ImportCustomObjects
// Description: Demonstrates how to convert hierarchical Order, Customer, and Item models into a flat list using LINQ SelectMany, define a custom column order, and write the data to an Excel worksheet with header rows, date formatting, and automatic numeric conversion via Cells.ImportCustomObjects. The workbook is saved as FlattenedOrders.xlsx.
// Keywords: Aspose.Cells | ImportCustomObjects | C# Excel export | flatten nested collections | .NET data export to Excel | LINQ SelectMany | custom column mapping | Excel workbook generation | order report automation | Excel automation .NET
// Common Searches: flatten nested C# objects for Excel using Aspose.Cells | ImportCustomObjects custom column order date format | export order and item list to Excel with Aspose.Cells | convert hierarchical data to flat rows for Excel in .NET | Aspose.Cells import list of objects with header row
// Developer Intent: The developer needs to transform hierarchical order data into a flat structure and write it to an Excel file using Aspose.Cells' ImportCustomObjects method.
// Use Cases: Create an order‑by‑item spreadsheet where each product line appears on its own row with customer and order details. | Generate a sales export file for ERP or accounting systems that requires flat rows and numeric/date formatting. | Produce a daily transaction report with ready‑to‑analyze data, including formatted dates and numeric values.
// AI Prompts: Show how to apply a custom currency format to the Price column in the ImportCustomObjects call. | Provide code that adds a summary row calculating total Quantity and total sales amount after the import. | Explain strategies for handling nullable fields when flattening nested objects for ImportCustomObjects.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsCustomImportDemo
{
    // Original nested data models
    // Demonstrates how to convert hierarchical Order, Customer, and Item models into a flat list using LINQ SelectMany, define a custom column order, and write the data to an Excel worksheet with header rows, date formatting, and automatic numeric conversion via Cells.ImportCustomObjects. The workbook is saved as FlattenedOrders.xlsx.
    public class Order
    {
        public int OrderId { get; set; }
        public Customer Customer { get; set; } = null!;
        public List<Item> Items { get; set; } = new();
        public DateTime OrderDate { get; set; }
    }

    public class Customer
    {
        public string Name { get; set; } = null!;
        public string Email { get; set; } = null!;
    }

    public class Item
    {
        public string ProductName { get; set; } = null!;
        public int Quantity { get; set; }
        public decimal Price { get; set; }
    }

    // Flattened representation for import
    public class OrderFlat
    {
        public int OrderId { get; set; }
        public string CustomerName { get; set; } = null!;
        public string CustomerEmail { get; set; } = null!;
        public DateTime OrderDate { get; set; }
        public string ProductName { get; set; } = null!;
        public int Quantity { get; set; }
        public decimal Price { get; set; }
    }

    public static class Demo
    {
        public static void Run()
        {
            // Prepare sample nested data
            var orders = new List<Order>
            {
                new Order
                {
                    OrderId = 1001,
                    OrderDate = new DateTime(2024, 1, 15),
                    Customer = new Customer { Name = "Alice", Email = "alice@example.com" },
                    Items = new List<Item>
                    {
                        new Item { ProductName = "Laptop", Quantity = 1, Price = 1200.00m },
                        new Item { ProductName = "Mouse", Quantity = 2, Price = 25.50m }
                    }
                },
                new Order
                {
                    OrderId = 1002,
                    OrderDate = new DateTime(2024, 2, 3),
                    Customer = new Customer { Name = "Bob", Email = "bob@example.com" },
                    Items = new List<Item>
                    {
                        new Item { ProductName = "Keyboard", Quantity = 1, Price = 75.00m }
                    }
                }
            };

            // Flatten nested collections into a list of OrderFlat objects
            List<OrderFlat> flatList = orders
                .SelectMany(o => o.Items, (o, i) => new OrderFlat
                {
                    OrderId = o.OrderId,
                    OrderDate = o.OrderDate,
                    CustomerName = o.Customer.Name,
                    CustomerEmail = o.Customer.Email,
                    ProductName = i.ProductName,
                    Quantity = i.Quantity,
                    Price = i.Price
                })
                .ToList();

            // Define the order of columns to be imported
            string[] propertyNames = new[]
            {
                "OrderId",
                "OrderDate",
                "CustomerName",
                "CustomerEmail",
                "ProductName",
                "Quantity",
                "Price"
            };

            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import the flattened data into the worksheet
            cells.ImportCustomObjects(
                flatList,
                propertyNames,
                true,   // include header row
                0,      // start at first row (A1)
                0,      // start at first column
                flatList.Count,
                true,   // insert rows if needed
                "yyyy-MM-dd",
                true    // convert string to number where possible
            );

            // Save the workbook to a file
            workbook.Save("FlattenedOrders.xlsx");
        }
    }

    // Entry point for the console application
    public static class Program
    {
        public static void Main()
        {
            try
            {
                Demo.Run();
                Console.WriteLine("Workbook created successfully: FlattenedOrders.xlsx");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
