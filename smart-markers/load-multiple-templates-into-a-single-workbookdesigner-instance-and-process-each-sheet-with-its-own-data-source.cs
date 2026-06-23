using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsMultipleTemplatesDemo
{
    // Sample data classes
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public Person(string name, int age) { Name = name; Age = age; }
    }

    public class Product
    {
        public string ProductName { get; set; }
        public double Price { get; set; }
        public Product(string name, double price) { ProductName = name; Price = price; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load individual template workbooks
            Workbook template1 = new Workbook("Template1.xlsx"); // contains smart markers for Person
            Workbook template2 = new Workbook("Template2.xlsx"); // contains smart markers for Product

            // Create a master workbook and combine the templates into it
            Workbook masterWorkbook = new Workbook();
            masterWorkbook.Combine(template1);
            masterWorkbook.Combine(template2);

            // Initialize WorkbookDesigner with the combined workbook
            WorkbookDesigner designer = new WorkbookDesigner(masterWorkbook);

            // Prepare data sources for each sheet
            List<Person> persons = new List<Person>
            {
                new Person("John Doe", 30),
                new Person("Jane Smith", 28)
            };

            List<Product> products = new List<Product>
            {
                new Product("Laptop", 1200.50),
                new Product("Smartphone", 799.99)
            };

            // Process first sheet (index 0) with Person data source
            designer.ClearDataSource();                         // Ensure previous sources are cleared
            designer.SetDataSource("Person", persons);          // Bind data source name used in Template1
            designer.Process(0, true);                          // Process only sheet 0

            // Process second sheet (index 1) with Product data source
            designer.ClearDataSource();
            designer.SetDataSource("Product", products);        // Bind data source name used in Template2
            designer.Process(1, true);                          // Process only sheet 1

            // Save the final workbook
            masterWorkbook.Save("CombinedOutput.xlsx");
        }
    }
}