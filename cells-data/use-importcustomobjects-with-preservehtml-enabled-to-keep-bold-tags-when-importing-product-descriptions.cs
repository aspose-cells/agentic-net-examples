using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

public class Product
{
    public string Name { get; set; }
    public string Description { get; set; }
    public decimal Price { get; set; }
}

public class ImportCustomObjectsPreserveHtmlDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample product list with HTML bold tags in the description
        List<Product> products = new List<Product>
        {
            new Product { Name = "Apple", Description = "<b>Fresh</b> and juicy", Price = 2.99m },
            new Product { Name = "Orange", Description = "<b>Sweet</b> citrus fruit", Price = 1.99m }
        };

        // Set import options to treat values as HTML strings
        ImportTableOptions options = new ImportTableOptions
        {
            IsFieldNameShown = true,   // include property names as header row
            TotalRows = products.Count,
            InsertRows = true,
            IsHtmlString = true        // preserve HTML tags during import
        };

        // Import the custom objects starting at cell A1 (row 0, column 0)
        sheet.Cells.ImportCustomObjects((ICollection)products, 0, 0, options);

        // Save the workbook as an Excel file
        workbook.Save("Products.xlsx", SaveFormat.Xlsx);

        // Save the workbook as HTML while keeping the HTML tags in cells
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ParseHtmlTagInCell = false   // preserve HTML tags instead of parsing them
        };
        workbook.Save("Products.html", htmlOptions);
    }
}