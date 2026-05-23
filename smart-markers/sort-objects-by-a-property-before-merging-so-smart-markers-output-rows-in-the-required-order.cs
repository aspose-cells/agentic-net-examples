using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerSortingDemo
{
    // Sample data class
    public class Product
    {
        public int ProductID { get; set; }
        public string ProductName { get; set; }
        public decimal Price { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create unsorted list of products
            List<Product> products = new List<Product>
            {
                new Product { ProductID = 3, ProductName = "Product C", Price = 15.99m },
                new Product { ProductID = 1, ProductName = "Product A", Price = 10.50m },
                new Product { ProductID = 2, ProductName = "Product B", Price = 12.75m }
            };

            // Sort the list by ProductID (ascending) before merging
            products.Sort((x, y) => x.ProductID.CompareTo(y.ProductID));

            // Create a workbook and set up smart markers
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("&=$ProductID");
            sheet.Cells["B1"].PutValue("&=$ProductName");
            sheet.Cells["C1"].PutValue("&=$Price");

            // Initialize WorkbookDesigner with the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Assign the sorted data source
            designer.SetDataSource("Products", products);

            // Process smart markers
            designer.Process();

            // Save the result
            workbook.Save("SortedProductsOutput.xlsx");
        }
    }
}