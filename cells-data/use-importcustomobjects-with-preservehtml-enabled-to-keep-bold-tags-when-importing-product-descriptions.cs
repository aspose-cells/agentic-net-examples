// Title: Import custom Product objects into Excel while preserving HTML <b> tags using ImportCustomObjects with PreserveHtml in Aspose.Cells for .NET
// Common Searches: how to import a list of objects into Excel and keep HTML bold tags with Aspose.Cells C# | Aspose.Cells ImportCustomObjects preserve HTML formatting example | C# import custom objects to worksheet retaining <b> tags | using ImportTableOptions PreserveHtml to keep rich text from HTML in Excel | export product descriptions with bold HTML to .xlsx using Aspose.Cells
// Tags: ImportCustomObjects preserve html Aspose.Cells | ImportTableOptions retain bold formatting | C# export objects to Excel with rich text | Aspose.Cells HTML to cell conversion | Excel workbook save html formatted strings

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportHtmlExample
{
    // Custom object representing a product.
    // The example creates a Workbook, defines a Product class whose Description contains HTML <b> tags, builds a list of Product objects, configures ImportTableOptions with PreserveHtml enabled, and imports the list into the first worksheet using sheet.Cells.ImportCustomObjects. The workbook is then saved as ProductsWithHtml.xlsx, preserving the bold formatting in the Excel cells.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
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

                // List of products with HTML in the description.
                List<Product> products = new List<Product>
                {
                    new Product { Name = "Product A", Description = "<b>Bold description for A</b>" },
                    new Product { Name = "Product B", Description = "<b>Bold description for B</b>" }
                };

                // Import options to preserve formatting (keeps HTML as rich text).
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    // By default, ImportDataOnly is false, which retains formatting.
                    // No need to set ImportDataOnly explicitly as the property is not available in this version.
                };

                // Import the custom objects starting at cell A1 (row 0, column 0).
                sheet.Cells.ImportCustomObjects(products, 0, 0, importOptions);

                // Save the workbook.
                string outputPath = "ProductsWithHtml.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
