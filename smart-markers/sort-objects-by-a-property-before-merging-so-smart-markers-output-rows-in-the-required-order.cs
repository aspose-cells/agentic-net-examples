// Title: C# – Sort Data Before Processing Smart Markers with Aspose.Cells WorkbookDesigner
// Description: This example demonstrates how to load an Excel template that contains smart markers, sort a collection of Product objects by a chosen property (e.g., Price) using LINQ, bind the sorted list to a WorkbookDesigner, process the smart markers so rows are generated in the sorted order, and save the resulting workbook.
// Keywords: Aspose.Cells C# smart markers sorting | WorkbookDesigner bind sorted collection | LINQ OrderBy with Aspose.Cells | Excel template smart marker rows order | C# generate ordered Excel report
// Common Searches: sort list before Aspose.Cells smart markers C# | WorkbookDesigner generate rows in custom order | how to order smart marker output by price | C# example for sorting data used by smart markers
// Developer Intent: Prepare the data in the required sequence before binding it to WorkbookDesigner so that smart marker expansion follows that order.
// Use Cases: Create a product catalog Excel file with items listed from lowest to highest price. | Generate an invoice where line‑item rows appear sorted by amount. | Produce a sales report that lists transactions chronologically by sorting a Date field before smart marker processing.
// AI Prompts: Write C# code that loads an Excel template, sorts a list of objects by a specified field, binds the sorted list to Aspose.Cells WorkbookDesigner, processes smart markers, and saves the output file. | Explain how LINQ OrderBy can be combined with Aspose.Cells smart markers to control the order of generated rows. | Provide a step‑by‑step guide for sorting multiple related collections before processing smart markers with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace SmartMarkerSortingDemo
{
    // Sample data class
    // This example demonstrates how to load an Excel template that contains smart markers, sort a collection of Product objects by a chosen property (e.g., Price) using LINQ, bind the sorted list to a WorkbookDesigner, process the smart markers so rows are generated in the sorted order, and save the resulting workbook.
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
            // Load a workbook that contains smart markers (e.g., "&=$ProductID", "&=$ProductName", "&=$Price")
            Workbook workbook = new Workbook("Template.xlsx");

            // Prepare unsorted data
            List<Product> products = new List<Product>
            {
                new Product { ProductID = 3, ProductName = "Product C", Price = 15.99m },
                new Product { ProductID = 1, ProductName = "Product A", Price = 10.50m },
                new Product { ProductID = 2, ProductName = "Product B", Price = 12.75m }
            };

            // Sort the collection by the desired property (Price ascending)
            List<Product> sortedProducts = products.OrderBy(p => p.Price).ToList();

            // Create a WorkbookDesigner instance and bind the sorted data source
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Products", sortedProducts);

            // Process smart markers – rows will be generated in the order of the sorted collection
            designer.Process();

            // Save the result
            workbook.Save("SortedSmartMarkersOutput.xlsx");
        }
    }
}
