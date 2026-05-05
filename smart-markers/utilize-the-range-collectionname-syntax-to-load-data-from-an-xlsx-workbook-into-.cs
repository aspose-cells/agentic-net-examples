using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Sample data class that matches the fields in the Excel template
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public int Stock { get; set; }
    }

    public class LoadDataWithRangeSyntax
    {
        public static void Run()
        {
            // 1. Prepare a collection of data to be loaded into the workbook
            List<Product> products = new List<Product>
            {
                new Product { Name = "Laptop", Price = 1200.50, Stock = 15 },
                new Product { Name = "Smartphone", Price = 799.99, Stock = 30 },
                new Product { Name = "Tablet", Price = 450.75, Stock = 20 }
            };

            // 2. Create a workbook and add a smart marker {{RANGE Products}} to the first sheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("{{RANGE Products}}");

            // 3. Create a WorkbookDesigner to bind the collection to the smart marker
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // 4. Set the data source using the collection name that matches the {{RANGE}} marker
            designer.SetDataSource("Products", products);

            // 5. Process the smart markers – this will populate the worksheet with the collection data
            designer.Process();

            // 6. Save the resulting workbook
            workbook.Save("ResultWithProducts.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            LoadDataWithRangeSyntax.Run();
            Console.WriteLine("Workbook generated successfully.");
        }
    }
}