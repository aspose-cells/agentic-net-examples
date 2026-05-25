using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportCustomObjectsHtmlDemo
{
    // Custom object representing a product with an HTML‑formatted description
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;   // May contain <b>...</b> tags
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Prepare sample data – descriptions contain bold HTML tags
                List<Product> products = new List<Product>
                {
                    new Product
                    {
                        Name = "Apple",
                        Description = "<b>Fresh</b> and juicy apples",
                        Price = 2.99m
                    },
                    new Product
                    {
                        Name = "Orange",
                        Description = "Citrus fruit with <b>rich</b> vitamin C",
                        Price = 1.79m
                    }
                };

                // 3. Configure import options to treat values as HTML strings
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    // Show property names in the first row (optional)
                    IsFieldNameShown = true,
                    // Preserve HTML tags instead of parsing them
                    IsHtmlString = true,
                    // Number of rows to import (same as the collection count)
                    TotalRows = products.Count,
                    // Insert rows if needed
                    InsertRows = true
                };

                // 4. Import the custom objects starting at cell A1 (row 0, column 0)
                worksheet.Cells.ImportCustomObjects((ICollection)products, 0, 0, importOptions);

                // 5. Save the workbook to an XLSX file
                string outputPath = "ProductsWithHtml.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Import completed. Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}