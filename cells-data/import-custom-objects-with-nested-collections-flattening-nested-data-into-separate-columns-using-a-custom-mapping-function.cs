// Title: Flatten hierarchical Order objects and import them as rows into an Excel worksheet with Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates over a List<Order> containing OrderItem collections, creates a flat List<OrderFlat>, and uses Aspose.Cells Workbook.Worksheets[0].Cells.ImportCustomObjects to write the data to an XLSX file with headers. | Generate a C# mapping function that converts nested order data into a flat structure and shows how to configure ImportTableOptions (headers, numeric conversion, date format) for exporting to Excel with Aspose.Cells.
// Common Searches: asp.net flatten list of objects with child collections for aspose.cells importcustomobjects | c# export hierarchical order data to xlsx using aspose.cells ImportCustomObjects | how to map Order and OrderItem classes to a flat table for Excel export in Aspose.Cells | using ImportTableOptions to include headers when importing custom objects into Excel with Aspose.Cells | convert nested collections to flat rows before saving workbook in Aspose.Cells .NET
// Tags: flatten nested collections Aspose.Cells | ImportCustomObjects C# Excel export | hierarchical data to flat rows Aspose.Cells | order data export to XLSX Aspose.Cells | custom import options Excel Aspose.Cells

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

// The example defines Order and OrderItem classes, builds sample nested order data, flattens it into OrderFlat objects, and uses Aspose.Cells ImportCustomObjects with ImportTableOptions (headers, numeric conversion, date format) to write the flat rows to a new workbook saved as FlattenedOrders.xlsx.
public class Order
{
    public int OrderId { get; set; }
    public string Customer { get; set; } = string.Empty;
    public List<OrderItem> Items { get; set; } = new List<OrderItem>();
}

public class OrderItem
{
    public string Product { get; set; } = string.Empty;
    public int Quantity { get; set; }
}

// Flat representation for import
public class OrderFlat
{
    public int OrderId { get; set; }
    public string Customer { get; set; } = string.Empty;
    public string Product { get; set; } = string.Empty;
    public int Quantity { get; set; }
}

public class ImportNestedObjectsDemo
{
    public static void Main()
    {
        try
        {
            // Sample nested data
            var orders = new List<Order>
            {
                new Order
                {
                    OrderId = 1001,
                    Customer = "Alice",
                    Items = new List<OrderItem>
                    {
                        new OrderItem { Product = "Apple", Quantity = 10 },
                        new OrderItem { Product = "Banana", Quantity = 5 }
                    }
                },
                new Order
                {
                    OrderId = 1002,
                    Customer = "Bob",
                    Items = new List<OrderItem>
                    {
                        new OrderItem { Product = "Orange", Quantity = 7 }
                    }
                }
            };

            // Flatten nested collections
            var flatList = new List<OrderFlat>();
            foreach (var order in orders)
            {
                foreach (var item in order.Items)
                {
                    flatList.Add(new OrderFlat
                    {
                        OrderId = order.OrderId,
                        Customer = order.Customer,
                        Product = item.Product,
                        Quantity = item.Quantity
                    });
                }
            }

            // Create workbook and worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Configure import options
            var options = new ImportTableOptions
            {
                IsFieldNameShown = true,   // include property names as header row
                InsertRows = true,         // add rows if needed
                ConvertNumericData = true, // convert numeric strings
                DateFormat = "yyyy-MM-dd"
            };

            // Import the flattened data starting at cell A1 (row 0, column 0)
            sheet.Cells.ImportCustomObjects((ICollection)flatList, 0, 0, options);

            // Save the workbook
            string outputPath = "FlattenedOrders.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
