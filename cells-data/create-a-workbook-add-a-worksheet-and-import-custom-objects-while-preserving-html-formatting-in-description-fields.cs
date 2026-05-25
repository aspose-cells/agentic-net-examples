using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomObjectImport
{
    // Define a custom class with an HTML‑formatted description field
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public string Description { get; set; }   // HTML content, e.g., "<b>Best seller</b>"
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook (uses the Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet (the workbook already contains one)
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Prepare a collection of custom objects with HTML in the Description field
            List<Product> products = new List<Product>
            {
                new Product
                {
                    Name = "Apple",
                    Price = 2.99m,
                    Description = "<b>Fresh and juicy</b>"
                },
                new Product
                {
                    Name = "Orange",
                    Price = 1.99m,
                    Description = "<i>Sweet and tangy</i>"
                }
            };

            // 4. Convert the collection to an ICellsDataTable using the factory (rule: Workbook.CellsDataTableFactory)
            CellsDataTableFactory factory = workbook.CellsDataTableFactory;
            ICellsDataTable dataTable = factory.GetInstance(products);

            // 5. Set import options to treat cell values as HTML strings
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,   // optional: include property names as header row
                IsHtmlString = true        // preserve HTML formatting
            };

            // 6. Import the data table into the worksheet starting at cell A1 (row 0, column 0)
            sheet.Cells.ImportData(dataTable, 0, 0, importOptions);

            // 7. Save the workbook to a file (uses Workbook.Save(string) rule)
            workbook.Save("CustomObjectsWithHtml.xlsx");

            Console.WriteLine("Workbook created and saved successfully.");
        }
    }
}