using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerNestedDemo
{
    // Sample data classes with nested properties
    public class Order
    {
        public Customer Customer { get; set; }
        public List<Item> Items { get; set; }
        public DateTime OrderDate { get; set; }
    }

    public class Customer
    {
        public string Name { get; set; }
        public Address Address { get; set; }
    }

    public class Address
    {
        public string City { get; set; }
        public string Street { get; set; }
    }

    public class Item
    {
        public string Product { get; set; }
        public int Quantity { get; set; }
        public double Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers
            // Example smart markers placed in the template:
            //   A1: "&=$Order.Customer.Name"
            //   A2: "&=$Order.Customer.Address.City"
            //   A3: "&=$Order.OrderDate"
            //   A5: "&=$Order.Items"
            //   B5: "&=$Order.Items.Product"
            //   C5: "&=$Order.Items.Quantity"
            //   D5: "&=$Order.Items.Price"
            Workbook template = new Workbook("template.xlsx");

            // Prepare nested data source
            Order order = new Order
            {
                OrderDate = DateTime.Today,
                Customer = new Customer
                {
                    Name = "John Doe",
                    Address = new Address
                    {
                        City = "New York",
                        Street = "5th Avenue"
                    }
                },
                Items = new List<Item>
                {
                    new Item { Product = "Apple",  Quantity = 10, Price = 0.5 },
                    new Item { Product = "Banana", Quantity = 5,  Price = 0.3 },
                    new Item { Product = "Orange", Quantity = 8,  Price = 0.4 }
                }
            };

            // Initialize WorkbookDesigner and bind the data source using the variable name "Order"
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = template;
            designer.SetDataSource("Order", order);

            // Process smart markers (dot notation will resolve nested properties)
            designer.Process();

            // Save the populated workbook
            designer.Workbook.Save("output.xlsx");
        }
    }
}