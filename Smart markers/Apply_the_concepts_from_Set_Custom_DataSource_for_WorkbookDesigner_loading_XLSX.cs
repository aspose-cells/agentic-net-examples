using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomDataSourceDemo
{
    // Sample POCO class to be used as a custom data source
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }

        public Product(string name, double price)
        {
            Name = name;
            Price = price;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the template workbook (must contain smart markers like &Product.Name, &Product.Price)
            string templatePath = "Template.xlsx";
            Workbook workbook = new Workbook(templatePath); // uses Workbook(string) constructor

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook); // uses WorkbookDesigner(Workbook) constructor

            // Prepare custom data source: a list of Product objects
            List<Product> products = new List<Product>
            {
                new Product("Apple", 1.20),
                new Product("Banana", 0.80),
                new Product("Cherry", 2.50)
            };

            // Bind the custom data source to the smart marker name "Product"
            designer.SetDataSource("Product", products); // uses SetDataSource(string, object)

            // Process the smart markers and populate the worksheet with data
            designer.Process(); // processes all smart markers

            // Save the result workbook
            string outputPath = "Result.xlsx";
            designer.Workbook.Save(outputPath); // uses Workbook.Save(string)

            Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
        }
    }
}