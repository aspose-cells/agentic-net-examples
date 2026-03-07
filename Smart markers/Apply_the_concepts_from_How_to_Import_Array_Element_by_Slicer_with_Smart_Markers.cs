using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsSmartMarkersDemo
{
    // Sample data class representing an item that will be bound to smart markers
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public int Stock { get; set; }
        public DateTime ReleaseDate { get; set; }
    }

    public class ImportArrayBySlicerWithSmartMarkers
    {
        public static void Run()
        {
            // Load the template workbook that contains smart markers (e.g., &=$Products.Name)
            Workbook workbook = new Workbook("template.xlsx"); // <-- load rule

            // Prepare a list of products that will be used as the data source for the smart markers
            List<Product> products = new List<Product>
            {
                new Product { Name = "Laptop", Price = 999.99, Stock = 25, ReleaseDate = new DateTime(2023, 5, 1) },
                new Product { Name = "Smartphone", Price = 699.49, Stock = 40, ReleaseDate = new DateTime(2023, 6, 15) },
                new Product { Name = "Tablet", Price = 399.00, Stock = 30, ReleaseDate = new DateTime(2023, 7, 20) }
            };

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set the data source for the smart markers. The name "Products" must match the marker prefix.
            designer.SetDataSource("Products", products);

            // Process all smart markers in the workbook (including slicer‑linked markers)
            designer.Process(); // <-- process rule

            // Save the populated workbook
            workbook.Save("output.xlsx"); // <-- save rule
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            ImportArrayBySlicerWithSmartMarkers.Run();
            Console.WriteLine("Workbook generated successfully.");
        }
    }
}