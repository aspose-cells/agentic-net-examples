// Title: Preserve HTML bold tags while importing custom objects into Excel using Aspose.Cells for .NET
// Common Searches: how to keep HTML bold tags when using ImportCustomObjects in Aspose.Cells C# | Aspose.Cells import custom objects with HTML formatting in .NET | preserve <b> tags in Excel cells after importing data with Aspose.Cells | C# example of ImportCustomObjects preserving HTML tags in a worksheet | import product description containing HTML into Excel using Aspose.Cells
// Tags: ImportCustomObjects HTML formatting Aspose.Cells | preserve bold tags Excel import C# | product description HTML to Excel Aspose | custom object array import Aspose.Cells .NET | worksheet cells import HTML content C#

using System;
using Aspose.Cells;

namespace AsposeCellsImportHtmlExample
{
    // Define a custom class representing a product.
    // The example defines a Product class whose Description may contain HTML <b> tags, creates an array of Product objects, converts it to a two‑dimensional object array, and uses Worksheet.Cells.ImportCustomObjects with default options to import the data into the first worksheet while preserving the HTML bold formatting, then saves the workbook as ProductsWithHtml.xlsx.
    public class Product
    {
        // The description may contain HTML tags such as <b> for bold text.
        public string? Description { get; set; }

        // Additional fields can be added as needed.
        public string? Name { get; set; }
        public double Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Prepare sample data with HTML bold tags in the description.
                Product[] products = new Product[]
                {
                    new Product { Name = "Laptop", Price = 999.99, Description = "<b>High‑performance</b> laptop with 16GB RAM" },
                    new Product { Name = "Smartphone", Price = 599.50, Description = "Latest model with <b>OLED</b> display" },
                    new Product { Name = "Headphones", Price = 199.00, Description = "Noise‑cancelling <b>over‑ear</b> headphones" }
                };

                // Convert the array of custom objects to an object array required by ImportCustomObjects.
                // Each row will contain the properties of the Product class in the order they are defined.
                object[,] data = new object[products.Length, 3];
                for (int i = 0; i < products.Length; i++)
                {
                    data[i, 0] = products[i].Name;
                    data[i, 1] = products[i].Price;
                    data[i, 2] = products[i].Description;
                }

                // Set import options (default options are sufficient for HTML handling).
                ImportTableOptions importOptions = new ImportTableOptions();

                // Import the data starting at cell A1 (row 0, column 0).
                sheet.Cells.ImportCustomObjects(data, 0, 0, importOptions);

                // Save the workbook to a file.
                string outputPath = "ProductsWithHtml.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
