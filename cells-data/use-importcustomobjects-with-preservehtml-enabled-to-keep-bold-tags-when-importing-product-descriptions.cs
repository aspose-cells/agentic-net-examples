// Title: Preserve HTML Tags When Importing Custom Objects with Aspose.Cells ImportCustomObjects (C#)
// Description: Shows how to import a List<Product> into an Excel worksheet using Aspose.Cells for .NET while retaining HTML formatting (e.g., <b>, <i>, <u>) by setting IsHtmlString in ImportTableOptions and saving the result as ProductsWithHtml.xlsx.
// Keywords: Aspose.Cells ImportCustomObjects | Preserve HTML tags in Excel | IsHtmlString true example | C# import custom objects with HTML | ImportTableOptions HTML preservation | Export product catalog to Excel | .NET Excel rich text import | Aspose.Cells keep bold italic | Excel cell HTML formatting | Aspose.Cells ImportCustomObjects C#
// Common Searches: Aspose.Cells keep HTML tags when importing | ImportCustomObjects preserve bold tags C# | IsHtmlString option Aspose.Cells | How to import list of objects with HTML formatting | Export HTML‑styled data to Excel using Aspose
// Developer Intent: Import a collection of objects into an Excel sheet while preserving embedded HTML markup such as bold, italic, and underline tags.
// Use Cases: Create a product catalog Excel file where descriptions contain rich‑text HTML from a database. | Generate price lists with styled text directly from C# models without post‑processing. | Transfer API‑returned HTML snippets into Excel for reporting while maintaining visual emphasis.
// AI Prompts: Write C# code that uses Aspose.Cells ImportCustomObjects with IsHtmlString enabled to keep HTML tags in cell values. | Explain how ImportTableOptions.IsHtmlString influences HTML preservation during custom object import. | Provide a step‑by‑step tutorial for importing a list of objects with <b>, <i>, and <u> tags into an Excel workbook using Aspose.Cells for .NET.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsHtmlImportDemo
{
    // Custom object representing a product with an HTML‑formatted description
    // Shows how to import a List<Product> into an Excel worksheet using Aspose.Cells for .NET while retaining HTML formatting (e.g., <b>, <i>, <u>) by setting IsHtmlString in ImportTableOptions and saving the result as ProductsWithHtml.xlsx.
    public class Product
    {
        public string Name { get; set; } = null!;
        public string Description { get; set; } = null!;   // May contain <b>, <i>, etc.
        public decimal Price { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data – description contains bold HTML tags
                List<Product> products = new List<Product>
                {
                    new Product
                    {
                        Name = "Apple",
                        Description = "<b>Fresh</b> and <i>crisp</i> apples",
                        Price = 2.99m
                    },
                    new Product
                    {
                        Name = "Orange",
                        Description = "<b>Juicy</b> oranges with <u>vitamin C</u>",
                        Price = 1.99m
                    }
                };

                // Configure import options:
                // - IsHtmlString = true tells Aspose.Cells that the string values may contain HTML.
                // - IsFieldNameShown = true writes the property names in the first row.
                // - InsertRows = true adds rows if needed.
                // - TotalRows = number of objects to import.
                ImportTableOptions options = new ImportTableOptions
                {
                    IsHtmlString = true,
                    IsFieldNameShown = true,
                    InsertRows = true,
                    TotalRows = products.Count,
                    ConvertNumericData = true,
                    DateFormat = "yyyy-MM-dd"
                };

                // Import the list of custom objects starting at cell A1 (row 0, column 0)
                sheet.Cells.ImportCustomObjects((ICollection)products, 0, 0, options);

                // Save the workbook (Excel format) – the HTML tags are preserved in the cells.
                workbook.Save("ProductsWithHtml.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
