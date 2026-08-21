// Title: Custom ICellsDataTable for Hierarchical Data with WorkbookDesigner (Aspose.Cells for .NET)
// Description: Demonstrates how to implement a custom ICellsDataTable that flattens nested Order‑Customer‑Item objects, assign it to WorkbookDesigner, and use smart markers to generate an Excel workbook in C# with Aspose.Cells.
// Keywords: Aspose.Cells | WorkbookDesigner | ICellsDataTable | custom data source | hierarchical data | smart markers | C# | .NET | Excel export | flatten nested collections | order items example
// Common Searches: Aspose.Cells custom ICellsDataTable example | WorkbookDesigner hierarchical data source C# | smart markers with nested objects Aspose.Cells | flatten order items for Excel export | set custom data source on WorkbookDesigner
// Developer Intent: The developer needs to bind a complex hierarchical object model to WorkbookDesigner using a custom data source that can be consumed by smart markers.
// Use Cases: Generate invoices where each order’s line items are listed by flattening order‑item relationships. | Create sales reports from nested customer‑order‑item structures without modifying the original object model. | Export a product catalog with categories and sub‑items to Excel by providing a custom ICellsDataTable that maps the hierarchy to flat rows.
// AI Prompts: Write a C# class that implements ICellsDataTable to flatten a list of Order objects with nested Item collections for WorkbookDesigner. | Show the code to set a custom hierarchical data source on WorkbookDesigner and process smart markers to produce an Excel file. | Explain how to extend the OrderDataSource to include additional fields such as OrderDate and ItemPrice while keeping the smart marker syntax unchanged.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomDataSourceDemo
{
    // Sample hierarchical data classes
    // Demonstrates how to implement a custom ICellsDataTable that flattens nested Order‑Customer‑Item objects, assign it to WorkbookDesigner, and use smart markers to generate an Excel workbook in C# with Aspose.Cells.
    public class Order
    {
        public int OrderId { get; set; }
        public Customer Customer { get; set; }
        public List<Item> Items { get; set; }
    }

    public class Customer
    {
        public string Name { get; set; }
    }

    public class Item
    {
        public string Product { get; set; }
        public int Quantity { get; set; }
    }

    // Custom ICellsDataTable implementation that flattens the hierarchical Order data
    public class OrderDataSource : ICellsDataTable
    {
        private readonly List<Order> _orders;
        private readonly List<object[]> _flattenedRows = new List<object[]>();
        private int _currentRow = -1;

        public OrderDataSource(List<Order> orders)
        {
            _orders = orders ?? new List<Order>();
            FlattenData();
        }

        // Convert hierarchical data into a flat table:
        // Columns: OrderId, CustomerName, Product, Quantity
        private void FlattenData()
        {
            foreach (var order in _orders)
            {
                if (order?.Items == null) continue;

                foreach (var item in order.Items)
                {
                    _flattenedRows.Add(new object[]
                    {
                        order.OrderId,
                        order.Customer?.Name,
                        item.Product,
                        item.Quantity
                    });
                }
            }
        }

        // Indexer by row and column index (safe bounds checking)
        public object this[int rowIndex, int columnIndex]
        {
            get
            {
                if (rowIndex >= 0 && rowIndex < _flattenedRows.Count &&
                    columnIndex >= 0 && columnIndex < Columns.Length)
                {
                    return _flattenedRows[rowIndex][columnIndex];
                }
                return null;
            }
        }

        // Indexer by row index (returns the whole row object, safe bounds checking)
        public object this[int rowIndex]
        {
            get
            {
                return (rowIndex >= 0 && rowIndex < _flattenedRows.Count) ? _flattenedRows[rowIndex] : null;
            }
        }

        // Indexer by column name
        public object this[string columnName]
        {
            get
            {
                int col = Array.IndexOf(Columns, columnName);
                return (col >= 0 && _currentRow >= 0 && _currentRow < _flattenedRows.Count)
                    ? _flattenedRows[_currentRow][col]
                    : null;
            }
        }

        public int RowCount => _flattenedRows.Count;
        public int ColumnCount => Columns.Length;
        public int Count => _flattenedRows.Count;

        public string[] Columns => new[] { "OrderId", "CustomerName", "Product", "Quantity" };

        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        public bool Next()
        {
            _currentRow++;
            return _currentRow < _flattenedRows.Count;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare hierarchical sample data
                var orders = new List<Order>
                {
                    new Order
                    {
                        OrderId = 1001,
                        Customer = new Customer { Name = "Alice" },
                        Items = new List<Item>
                        {
                            new Item { Product = "Laptop", Quantity = 1 },
                            new Item { Product = "Mouse", Quantity = 2 }
                        }
                    },
                    new Order
                    {
                        OrderId = 1002,
                        Customer = new Customer { Name = "Bob" },
                        Items = new List<Item>
                        {
                            new Item { Product = "Keyboard", Quantity = 1 }
                        }
                    }
                };

                // Create a new workbook and add smart markers
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header row
                sheet.Cells["A1"].PutValue("Order ID");
                sheet.Cells["B1"].PutValue("Customer");
                sheet.Cells["C1"].PutValue("Product");
                sheet.Cells["D1"].PutValue("Quantity");

                // Data row with smart markers (the same row will be repeated for each record)
                sheet.Cells["A2"].PutValue("&=Order.OrderId");
                sheet.Cells["B2"].PutValue("&=Order.CustomerName");
                sheet.Cells["C2"].PutValue("&=Order.Product");
                sheet.Cells["D2"].PutValue("&=Order.Quantity");

                // Initialize WorkbookDesigner with the workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Set the custom hierarchical data source using ICellsDataTable
                designer.SetDataSource("Order", new OrderDataSource(orders));

                // Process smart markers and populate data
                designer.Process();

                // Save the result
                string outputPath = "CustomHierarchicalDataOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
