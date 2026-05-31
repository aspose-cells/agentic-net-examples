using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public decimal Price { get; set; }
        public int Stock { get; set; }
        public DateTime Date { get; set; }
    }

    public class ImportCustomObjectsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Prepare a list of custom objects containing a DateTime property
                List<Product> products = new List<Product>
                {
                    new Product { Name = "Apple",  Price = 2.99m, Stock = 150, Date = new DateTime(2023, 12, 31) },
                    new Product { Name = "Orange", Price = 1.99m, Stock = 200, Date = new DateTime(2024, 1, 15) }
                };

                // Define the order of properties to be imported
                string[] propertyNames = { "Name", "Price", "Stock", "Date" };

                // Import the custom objects into the worksheet.
                // The dateFormatString parameter sets the desired display format for DateTime values.
                worksheet.Cells.ImportCustomObjects(
                    products,                // list of objects
                    propertyNames,           // property names to import
                    true,                    // show property names as header
                    0,                       // first row index
                    0,                       // first column index
                    products.Count,          // number of rows to import
                    true,                    // insert rows if needed
                    "dd-MMM-yyyy",           // custom date format
                    true                     // convert strings to numbers where possible
                );

                // Save the workbook to a file
                string outputPath = "CustomObjectsDateFormat.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ImportCustomObjectsDemo.Run();
        }
    }
}