using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomObjectImport
{
    // Sample custom object with an HTML formatted description
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public string Description { get; set; } // May contain HTML tags
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet (already added by default)
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare a list of custom objects
            List<Product> products = new List<Product>
            {
                new Product
                {
                    Name = "Apple",
                    Price = 2.99m,
                    Description = "<b>Fresh</b> and <i>crisp</i> apples."
                },
                new Product
                {
                    Name = "Orange",
                    Price = 1.99m,
                    Description = "<u>Juicy</u> oranges with <span style=\"color:orange;\">vibrant color</span>."
                }
            };

            // Obtain the factory for building ICellsDataTable from custom objects (rule usage)
            CellsDataTableFactory factory = workbook.CellsDataTableFactory;

            // Create a data table from the custom object list
            ICellsDataTable dataTable = factory.GetInstance(products);

            // Configure import options to treat string values as HTML
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,   // Include property names as header row
                IsHtmlString = true        // Preserve HTML formatting
            };

            // Import the data table into the worksheet starting at cell A1
            worksheet.Cells.ImportData(dataTable, 0, 0, importOptions);

            // Save the workbook to a file (lifecycle rule)
            workbook.Save("CustomObjectsWithHtml.xlsx");
        }
    }
}