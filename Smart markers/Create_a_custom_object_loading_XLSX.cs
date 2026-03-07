using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomObjectDemo
{
    // Define a custom data class
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Stock { get; set; }
        public DateTime ReleaseDate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            List<Product> products = new List<Product>
            {
                new Product { Name = "Laptop", Price = 1299.99m, Stock = 50, ReleaseDate = new DateTime(2023, 11, 15) },
                new Product { Name = "Smartphone", Price = 799.49m, Stock = 120, ReleaseDate = new DateTime(2024, 1, 5) },
                new Product { Name = "Tablet", Price = 499.00m, Stock = 80, ReleaseDate = new DateTime(2023, 9, 30) }
            };

            // Property names to import (order matters)
            string[] propertyNames = { "Name", "Price", "Stock", "ReleaseDate" };

            // Load an existing XLSX file (template) – replace with your actual file path
            string templatePath = "Template.xlsx";
            Workbook workbook = new Workbook(templatePath); // uses Workbook(string) constructor

            // Get the first worksheet where data will be imported
            Worksheet sheet = workbook.Worksheets[0];

            // Import the custom objects into the worksheet starting at cell A1 (row 0, column 0)
            // Parameters:
            //   list               : collection of custom objects
            //   propertyNames      : columns to import
            //   isPropertyNameShown: include header row
            //   firstRow           : start row index (0‑based)
            //   firstColumn        : start column index (0‑based)
            //   rowNumber          : number of rows to import
            //   insertRows         : insert rows if needed
            //   dateFormatString   : format for DateTime values
            //   convertStringToNumber : try converting strings to numbers
            int importedRows = sheet.Cells.ImportCustomObjects(
                products,
                propertyNames,
                true,          // show header
                0,             // first row
                0,             // first column
                products.Count,
                true,          // insert rows if needed
                "yyyy-MM-dd",  // date format
                true           // convert strings to numbers when possible
            );

            Console.WriteLine($"Imported {importedRows} rows into the worksheet.");

            // Save the workbook with the imported data
            string outputPath = "ProductsOutput.xlsx";
            workbook.Save(outputPath); // uses Workbook.Save(string) method

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}