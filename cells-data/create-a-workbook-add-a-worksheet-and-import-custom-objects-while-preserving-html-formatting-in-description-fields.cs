// Title: Import Custom Objects into an Aspose.Cells Worksheet while Preserving HTML Formatting
// Description: Creates a workbook, builds a CellsDataTable from a List of Product objects that contain HTML in the Description property, imports the table with ImportTableOptions.IsHtmlString enabled, and saves the file as CustomObjectsWithHtml.xlsx.
// Keywords: Aspose.Cells | Import custom objects | HTML formatting in Excel | ImportTableOptions IsHtmlString | CellsDataTableFactory C# | export objects to Excel | rich text Excel cells
// Common Searches: Aspose.Cells import list of objects with HTML tags | ImportData IsHtmlString true example | CellsDataTableFactory export HTML description | C# preserve HTML formatting in Excel cells using Aspose | how to keep bold italic colors when writing to Excel with Aspose.Cells
// Developer Intent: The developer needs to generate an Excel workbook, add a worksheet, and import a collection of custom objects so that any HTML markup in string fields is rendered as rich text in the cells.
// Use Cases: Product catalog where descriptions include bold, italic, or colored text defined in HTML. | Report that stores notes as HTML snippets and requires the formatting to appear in Excel. | Invoice sheet with item details that use HTML for branding or emphasis.
// AI Prompts: Show how to map object properties to custom column headers when using CellsDataTableFactory. | Give an example of importing both HTML‑formatted and plain‑text fields and auto‑adjusting column widths. | Explain how to exclude or rename specific properties during import with IsHtmlString enabled.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomObjectImport
{
    // Custom object with an HTML formatted description
    // Creates a workbook, builds a CellsDataTable from a List of Product objects that contain HTML in the Description property, imports the table with ImportTableOptions.IsHtmlString enabled, and saves the file as CustomObjectsWithHtml.xlsx.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public string Description { get; set; } // HTML content
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (adds automatically)
            Worksheet worksheet = workbook.Worksheets[0];

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
                    Description = "<span style=\"color:orange;\">Juicy</span> oranges."
                }
            };

            // Build a data table from the custom objects
            CellsDataTableFactory factory = workbook.CellsDataTableFactory;
            ICellsDataTable dataTable = factory.GetInstance(products);

            // Set import options to treat values as HTML strings
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,   // Include property names as header row
                IsHtmlString = true        // Preserve HTML formatting
            };

            // Import the data table into the worksheet starting at cell A1
            worksheet.Cells.ImportData(dataTable, 0, 0, importOptions);

            // Save the workbook
            workbook.Save("CustomObjectsWithHtml.xlsx");
        }
    }
}
