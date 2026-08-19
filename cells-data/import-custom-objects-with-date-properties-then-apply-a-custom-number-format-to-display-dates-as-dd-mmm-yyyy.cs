// Title: C# – Import Custom Objects with DateTime and Apply "dd-MMM-yyyy" Format using Aspose.Cells
// Description: Demonstrates how to create a Workbook, define a Product class with a DateTime property, build a List<Product>, and use Worksheet.Cells.ImportCustomObjects to import the list while specifying a custom date format (dd-MMM-yyyy). The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | ImportCustomObjects | C# | .NET | custom object import | DateTime format | dd-MMM-yyyy | Excel export | worksheet cell formatting | save workbook
// Common Searches: Aspose.Cells import list of objects with date format | C# ImportCustomObjects custom date pattern | how to set date format when importing objects into Excel | apply dd-MMM-yyyy to DateTime cells Aspose.Cells | export product list to Excel with formatted dates
// Developer Intent: Import a collection of custom C# objects into an Excel worksheet and display their DateTime values using a custom "dd-MMM-yyyy" format.
// Use Cases: Generate a sales report where release dates appear as 31-Dec-2023, 15-Jan-2024, etc. | Create an inventory sheet that shows restock dates in a consistent, readable format. | Export an event schedule to Excel while ensuring all dates follow the same custom pattern.
// AI Prompts: Write C# code that uses Aspose.Cells to import a list of objects containing DateTime fields and format the date column as "dd-MMM-yyyy". | Explain the purpose of the dateFormatString parameter in ImportCustomObjects and how to adapt it for different locales. | Show how to extend the example to include time formatting, e.g., "dd-MMM-yyyy HH:mm", for a DateTime column.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom object with a DateTime property
    // Demonstrates how to create a Workbook, define a Product class with a DateTime property, build a List<Product>, and use Worksheet.Cells.ImportCustomObjects to import the list while specifying a custom date format (dd-MMM-yyyy). The workbook is saved as an XLSX file.
    public class Product
    {
        public string Name { get; set; } = string.Empty;   // initialize to avoid nullable warning
        public decimal Price { get; set; }
        public int Stock { get; set; }
        public DateTime Date { get; set; }
    }

    public class ImportCustomObjectsWithDateFormatDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Prepare sample data
                List<Product> products = new List<Product>
                {
                    new Product { Name = "Apple", Price = 2.99m, Stock = 150, Date = new DateTime(2023, 12, 31) },
                    new Product { Name = "Orange", Price = 1.99m, Stock = 200, Date = new DateTime(2024, 1, 15) }
                };

                // Define the order of properties to import
                string[] propertyNames = { "Name", "Price", "Stock", "Date" };

                // Import the custom objects.
                // The dateFormatString parameter sets the desired display format for DateTime cells.
                int importedRows = worksheet.Cells.ImportCustomObjects(
                    products,                // ICollection list
                    propertyNames,           // property names to import
                    true,                    // show property names in the first row
                    0,                       // first row index
                    0,                       // first column index
                    products.Count,          // number of rows to import
                    true,                    // insert rows if needed
                    "dd-MMM-yyyy",           // custom date format
                    true                     // convert strings to numbers where possible
                );

                Console.WriteLine($"Successfully imported {importedRows} rows with custom date format.");

                // Save the workbook
                string outputPath = "ImportCustomObjectsWithDateFormat.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ImportCustomObjectsWithDateFormatDemo.Run();
        }
    }
}
