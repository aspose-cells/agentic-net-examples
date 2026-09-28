// Title: Use Aspose.Cells WorkbookDesigner to bind a List<Product> custom collection to smart markers and export to Excel in C#
// AI Prompts: Write C# code that creates a WorkbookDesigner, adds smart marker expressions for a product list, binds a List<Product> via SetDataSource, processes the markers, and saves the workbook as an .xlsx file. | Show how to insert header cells and smart marker placeholders (e.g., &=Products.Name) into a worksheet before attaching a custom object collection in Aspose.Cells. | Demonstrate proper error handling while populating an Excel template from a collection of objects using Aspose.Cells WorkbookDesigner in .NET.
// Common Searches: aspnet bind List<Product> to Aspose.Cells WorkbookDesigner smart markers | c# generate Excel rows from custom object collection using Aspose.Cells | how to use SetDataSource with a List of objects in Aspose.Cells WorkbookDesigner | populate Excel template with product data using smart markers in C# | Aspose.Cells WorkbookDesigner example for dynamic table generation from objects
// Tags: WorkbookDesigner SetDataSource with List<T> | Aspose.Cells smart markers collection binding | C# generate Excel from object list | dynamic row insertion using smart markers | export custom collection to .xlsx with Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // The sample creates a new Workbook, writes header cells, places smart marker expressions referencing a "Products" collection, builds a List<Product> with sample data, binds this collection to WorkbookDesigner via SetDataSource, processes the markers to fill the worksheet, and saves the result as CustomCollectionOutput.xlsx.
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public int Quantity { get; set; }

        public Product(string name, double price, int quantity)
        {
            Name = name;
            Price = price;
            Quantity = quantity;
        }
    }

    public class WorkbookDesignerCustomCollectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header row
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["C1"].PutValue("Quantity");

                // Smart markers that reference the collection named "Products"
                sheet.Cells["A2"].PutValue("&=Products.Name");
                sheet.Cells["B2"].PutValue("&=Products.Price");
                sheet.Cells["C2"].PutValue("&=Products.Quantity");

                // Prepare a custom collection of objects
                List<Product> products = new List<Product>
                {
                    new Product("Apple", 1.20, 50),
                    new Product("Banana", 0.80, 100),
                    new Product("Orange", 1.50, 75)
                };

                // Initialize the designer and assign the workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Bind the custom collection to the smart marker variable "Products"
                designer.SetDataSource("Products", products);

                // Process the smart markers to populate the worksheet
                designer.Process();

                // Save the populated workbook
                workbook.Save("CustomCollectionOutput.xlsx");
                Console.WriteLine("Workbook saved as CustomCollectionOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookDesignerCustomCollectionDemo.Run();
        }
    }
}
