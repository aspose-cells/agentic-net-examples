// Title: Create an Excel workbook with Aspose.Cells, import custom C# objects, and display HTML-formatted descriptions
// Common Searches: how to import a list of C# objects into Aspose.Cells without column headers | preserve HTML tags in Excel cells when using ImportCustomObjects in Aspose.Cells | set HtmlString property for cells after importing data with Aspose.Cells | Aspose.Cells import custom objects and render HTML in description column | C# Aspose.Cells create worksheet and import objects with HTML formatting
// Tags: ImportCustomObjects without field names | HtmlString rendering for cell values | create workbook and add worksheet Aspose.Cells | import C# objects with HTML description | preserve HTML formatting in Excel cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImportHtmlExample
{
    // Custom object with an HTML formatted description
    // The sample creates a new Workbook, adds a worksheet named "Products", builds a List<Product> where each Description contains HTML markup, and imports the list into the sheet starting at A1 using ImportCustomObjects with IsFieldNameShown set to false to omit property names as headers. After import it locates the description cells, assigns the raw HTML string to each cell's HtmlString property so the markup is rendered in Excel, and finally saves the file as "ProductsWithHtml.xlsx" with basic error handling.
    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Add a new worksheet (optional, the default workbook already contains one)
                Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet.Name = "Products";

                // 3. Prepare a list of custom objects
                List<Product> products = new List<Product>
                {
                    new Product
                    {
                        Id = 1,
                        Name = "Laptop",
                        Description = "<b>Powerful</b> laptop with <i>16GB RAM</i> and <u>SSD</u>."
                    },
                    new Product
                    {
                        Id = 2,
                        Name = "Smartphone",
                        Description = "Latest <font color='red'>Android</font> phone with <span style='color:blue'>high‑resolution</span> display."
                    },
                    new Product
                    {
                        Id = 3,
                        Name = "Headphones",
                        Description = "Noise‑cancelling <strong>over‑ear</strong> headphones."
                    }
                };

                // 4. Import the custom objects into the worksheet starting at cell A1
                //    Use ImportTableOptions to disable column name import.
                var importOptions = new ImportTableOptions
                {
                    IsFieldNameShown = false // do not import property names as column headers
                };
                sheet.Cells.ImportCustomObjects(products, 0, 0, importOptions);

                // 5. Preserve HTML formatting for the Description column (column C, index 2)
                //    After import, each description cell contains the raw HTML string.
                //    Setting the HtmlString property makes the cell render the HTML.
                int startRow = sheet.Cells.MaxDataRow - products.Count + 1; // first data row index
                for (int i = 0; i < products.Count; i++)
                {
                    try
                    {
                        int rowIndex = startRow + i;               // Row where the product is placed
                        Cell descCell = sheet.Cells[rowIndex, 2]; // Column C (index 2)

                        // Render the HTML stored in the cell
                        descCell.HtmlString = descCell.StringValue;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing row {i}: {ex.Message}");
                    }
                }

                // 6. Save the workbook to a file
                string outputPath = "ProductsWithHtml.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
