using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a template workbook with smart markers for nested data
        // ------------------------------------------------------------
        Workbook template = new Workbook();
        Worksheet sheet = template.Worksheets[0];

        // Header row
        sheet.Cells["A1"].PutValue("Customer Name");
        sheet.Cells["B1"].PutValue("Order ID");
        sheet.Cells["C1"].PutValue("Product");
        sheet.Cells["D1"].PutValue("Quantity");

        // Smart marker row – note the use of nested object paths
        sheet.Cells["A2"].PutValue("&=Customers.Name");
        sheet.Cells["B2"].PutValue("&=Customers.Orders.OrderID");
        sheet.Cells["C2"].PutValue("&=Customers.Orders.Product");
        sheet.Cells["D2"].PutValue("&=Customers.Orders.Quantity");

        // Define the range that contains the smart markers
        sheet.Cells.CreateRange("A2:D2").Name = "_CellsSmartMarkers";

        // ------------------------------------------------------------
        // 2. Save the template to a memory stream (simulating a file)
        // ------------------------------------------------------------
        using (MemoryStream templateStream = new MemoryStream())
        {
            template.Save(templateStream, SaveFormat.Xlsx);
            templateStream.Position = 0;

            // ------------------------------------------------------------
            // 3. Load the workbook into memory with optimized LoadOptions
            // ------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions
            {
                // Unparsed data is not needed for this scenario – improves performance
                KeepUnparsedData = false
            };
            Workbook workbook = new Workbook(templateStream, loadOptions);

            // ------------------------------------------------------------
            // 4. Prepare nested data (customers with a list of orders)
            // ------------------------------------------------------------
            var customers = new List<Customer>
            {
                new Customer
                {
                    Name = "Acme Corp",
                    Orders = new List<Order>
                    {
                        new Order { OrderID = 1001, Product = "Widget",      Quantity = 10 },
                        new Order { OrderID = 1002, Product = "Gadget",      Quantity = 5  }
                    }
                },
                new Customer
                {
                    Name = "Beta Ltd",
                    Orders = new List<Order>
                    {
                        new Order { OrderID = 2001, Product = "Thingamajig", Quantity = 7 }
                    }
                }
            };

            // ------------------------------------------------------------
            // 5. Bind the data source to the designer and process smart markers
            // ------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Customers", customers);
            designer.Process(); // processes the range named "_CellsSmartMarkers"

            // ------------------------------------------------------------
            // 6. Save the populated workbook
            // ------------------------------------------------------------
            workbook.Save("NestedSmartMarkersOutput.xlsx");
        }
    }

    // ----------------------------------------------------------------
    // Data model: Customer contains a collection of Order objects
    // ----------------------------------------------------------------
    public class Customer
    {
        public string Name { get; set; }
        public List<Order> Orders { get; set; }
    }

    public class Order
    {
        public int OrderID { get; set; }
        public string Product { get; set; }
        public int Quantity { get; set; }
    }
}