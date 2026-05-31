using System;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsSlicerImportDemo
{
    // Simple data model representing an order
    public class Order
    {
        public int Quantity { get; set; }
        public string Product { get; set; }

        public Order(string product, int quantity)
        {
            Product = product;
            Quantity = quantity;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data: an array of Order objects
            Order[] Orders = new Order[]
            {
                new Order("Apple", 10),   // index 0
                new Order("Banana", 20),  // index 1
                new Order("Orange", 30),  // index 2
                new Order("Grape", 40),   // index 3
                new Order("Mango", 50)    // index 4
            };

            // Use C# range syntax to get a subset (indexes 1 to 3 inclusive)
            // This corresponds to Orders[1], Orders[2], Orders[3]
            Order[] subset = Orders[1..4]; // end index is exclusive, so 4 gives up to 3

            // Extract the Quantity values from the subset
            int[] quantities = subset.Select(o => o.Quantity).ToArray();

            // Import the quantities vertically starting at cell B2 (row index 1, column index 1)
            // The 'true' flag indicates vertical orientation
            worksheet.Cells.ImportArray(quantities, firstRow: 1, firstColumn: 1, isVertical: true);

            // Optional: write a header for clarity
            worksheet.Cells[0, 1].PutValue("Quantity");

            // Save the workbook
            workbook.Save("SubsetImportDemo.xlsx");
        }
    }
}