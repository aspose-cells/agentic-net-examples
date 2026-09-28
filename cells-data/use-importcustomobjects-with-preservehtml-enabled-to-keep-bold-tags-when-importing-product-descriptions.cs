// Title: Import a List of objects into an Excel worksheet while preserving <b> HTML tags using Aspose.Cells ImportCustomObjects in C#
// AI Prompts: Generate C# code that uses Aspose.Cells ImportCustomObjects to load a List<Product> into a worksheet and keep HTML tags such as <b> intact in the cells. | Show how to set HtmlSaveOptions so that the workbook is saved to HTML without parsing the HTML tags stored in cells. | Demonstrate configuring ImportTableOptions to add a header row and enable automatic numeric conversion when importing custom objects.
// Common Searches: Aspose.Cells C# import custom objects preserve HTML formatting in cells | ImportCustomObjects keep bold tags when exporting to HTML | disable HTML tag parsing in Aspose.Cells HTML export | example of ImportTableOptions IsHtmlString true with List<T>
// Tags: ImportCustomObjects with IsHtmlString option | preserve HTML tags in Excel cells Aspose.Cells | HtmlSaveOptions ParseHtmlTagInCell false | import List<T> to worksheet C# Aspose.Cells | export workbook to HTML without tag parsing

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportHtmlDemo
{
    // Simple product class with an HTML‑formatted description
    // Demonstrates importing a List<Product> into a worksheet using ImportCustomObjects with IsHtmlString enabled to retain <b> tags, then saving the workbook as XLSX and as HTML with ParseHtmlTagInCell set to false to keep the tags unchanged.
    public class Product
    {
        public string Description { get; set; } = string.Empty;
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data – the description contains <b> tags
                List<Product> products = new List<Product>
                {
                    new Product { Description = "<b>Apple</b> – Fresh fruit", Price = 2.99m },
                    new Product { Description = "<b>Orange</b> – Citrus delight", Price = 1.79m }
                };

                // Configure import options to treat cell values as HTML strings
                ImportTableOptions options = new ImportTableOptions
                {
                    IsFieldNameShown = true,      // import property names as header row
                    IsHtmlString = true,          // preserve HTML tags (e.g., <b>)
                    TotalRows = products.Count,   // number of rows to import
                    InsertRows = true,            // insert rows if needed
                    ConvertNumericData = true,    // convert numeric values automatically
                    DateFormat = "yyyy-MM-dd"     // required by the class, not used here
                };

                // Import the collection into the worksheet starting at cell A1 (row 0, column 0)
                sheet.Cells.ImportCustomObjects((ICollection)products, 0, 0, options);

                // Save as Excel file (HTML tags are stored in the cell)
                workbook.Save("Products.xlsx", SaveFormat.Xlsx);

                // When exporting to HTML we want the tags to remain unchanged,
                // so disable parsing of HTML tags in cells.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ParseHtmlTagInCell = false
                };
                workbook.Save("Products.html", htmlOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
