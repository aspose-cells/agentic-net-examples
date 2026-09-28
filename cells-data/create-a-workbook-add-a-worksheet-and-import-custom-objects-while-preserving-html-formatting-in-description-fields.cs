// Title: Import a List of Product objects into an Aspose.Cells worksheet while preserving HTML tags in description cells (C#)
// AI Prompts: Generate C# code that builds an ICellsDataTable from a List<Product> and imports it into a worksheet with headers and HTML strings retained. | Show how to configure ImportTableOptions.IsHtmlString = true when calling Worksheet.Cells.ImportData to keep HTML formatting. | Outline the steps to save the workbook after importing custom objects that contain HTML‑formatted description fields.
// Common Searches: how to import a List<T> into Aspose.Cells and keep HTML formatting in cells | Aspose.Cells C# preserve HTML tags when exporting a collection of objects | using ImportTableOptions to retain HTML content during Excel export | convert custom object list to ICellsDataTable and write to .xlsx with styled text
// Tags: import custom objects ICellsDataTable C# | keep html formatting during Excel import | use ImportTableOptions for html content | export product list to xlsx with styled description | worksheet cells import data with headers

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomObjectImport
{
    // Custom data class with an HTML formatted description field
    // The program creates a new Workbook, converts a List<Product> (where Description contains HTML) into an ICellsDataTable, sets ImportTableOptions to show field names and treat values as HTML strings, imports the table starting at A1 of the first worksheet, and saves the file as CustomObjectsWithHtml.xlsx.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public string Description { get; set; }   // HTML content
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet (adds a worksheet automatically)
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data with HTML in the Description field
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
                    Description = "<u>Juicy</u> oranges with <span style='color:orange'>vibrant color</span>."
                }
            };

            // Build an ICellsDataTable from the custom object collection
            CellsDataTableFactory factory = workbook.CellsDataTableFactory;
            ICellsDataTable dataTable = factory.GetInstance(products);

            // Set import options to treat values as HTML strings
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,   // import property names as header row
                IsHtmlString = true        // preserve HTML formatting
            };

            // Import the data table into the worksheet starting at cell A1
            sheet.Cells.ImportData(dataTable, 0, 0, importOptions);

            // Save the workbook (lifecycle rule)
            workbook.Save("CustomObjectsWithHtml.xlsx");
        }
    }
}
