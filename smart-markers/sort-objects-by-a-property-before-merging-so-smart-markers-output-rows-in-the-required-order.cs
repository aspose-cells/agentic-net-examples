using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SmartMarkerSortingDemo
{
    // Sample data class
    public class Product
    {
        public int ProductID { get; set; }
        public string ProductName { get; set; }
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load a workbook that contains smart markers (e.g., &=$ProductID, &=$ProductName, &=$Price)
            Workbook workbook = new Workbook("Template.xlsx");

            // Create a WorkbookDesigner instance for processing smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare unsorted data
            List<Product> products = new List<Product>
            {
                new Product { ProductID = 3, ProductName = "Product C", Price = 15.99m },
                new Product { ProductID = 1, ProductName = "Product A", Price = 10.50m },
                new Product { ProductID = 2, ProductName = "Product B", Price = 12.75m }
            };

            // Sort the collection by the required property (ProductID) before merging
            List<Product> sortedProducts = products.OrderBy(p => p.ProductID).ToList();

            // Set the sorted collection as the data source for the smart markers
            designer.SetDataSource("Products", sortedProducts);

            // Process the smart markers and populate the worksheet
            designer.Process();

            // Save the resulting workbook
            workbook.Save("SortedSmartMarkersOutput.xlsx");
        }
    }
}