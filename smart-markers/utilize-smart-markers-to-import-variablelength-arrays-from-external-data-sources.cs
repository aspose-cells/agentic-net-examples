using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the template workbook that contains smart markers (e.g., &Products.Name, &Products.Price, &Products.Quantity)
        Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

        // Prepare a variable‑length collection as the external data source
        List<Product> products = new List<Product>
        {
            new Product { Name = "Apple",  Price = 1.20, Quantity = 10 },
            new Product { Name = "Banana", Price = 0.80, Quantity = 20 }
        };

        // Dynamically add more items to demonstrate variable length
        for (int i = 0; i < 5; i++)
        {
            products.Add(new Product
            {
                Name = $"Item{i}",
                Price = 0.5 + i,
                Quantity = i * 3
            });
        }

        // Initialize WorkbookDesigner with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Bind the collection to the smart marker name used in the template
        designer.SetDataSource("Products", products);

        // Process the smart markers; the range smart markers in the template will expand to fit the collection size
        designer.Process();

        // Save the populated workbook
        workbook.Save("Result.xlsx");
    }

    // Simple POCO class representing the data structure used by the smart markers
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public int Quantity { get; set; }
    }
}