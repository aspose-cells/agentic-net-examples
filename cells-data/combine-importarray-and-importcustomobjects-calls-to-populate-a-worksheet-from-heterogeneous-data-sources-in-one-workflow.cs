// Title: Combine ImportArray and ImportCustomObjects in Aspose.Cells for .NET – Populate Headers and Product List in One Worksheet
// Description: Shows how to use Aspose.Cells for .NET to import a string array of category headers horizontally and a List&lt;Product&gt; of custom objects vertically into the same worksheet, display property names, apply a custom date format, insert rows, convert strings to numbers, auto‑fit columns, and save the file as an XLSX document.
// Keywords: Aspose.Cells | ImportArray C# | ImportCustomObjects C# | .NET Excel export | populate worksheet from array and objects | header row import Aspose | custom object import Excel | date format Excel C# | auto fit columns Aspose.Cells | C# Excel automation
// Common Searches: Aspose.Cells import string array and object list in same sheet | ImportArray followed by ImportCustomObjects example C# | How to add headers and product data with Aspose.Cells | Combine heterogeneous data sources into Excel using Aspose.Cells | C# import custom objects with date formatting Aspose
// Developer Intent: The developer needs a single workflow that imports both a simple array of headers and a collection of custom objects into one Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Create an inventory report where category headers are placed across the top row and product details are listed below. | Generate a sales dashboard that mixes static label arrays with dynamic object collections in a single sheet. | Build an automated Excel export that combines configuration data (arrays) and runtime data (objects) with formatting and auto‑sizing.
// AI Prompts: Generate code to add a numeric array after the product rows using ImportArray. | Provide a snippet that applies conditional formatting to the imported product rows. | Explain how to safely handle null values in the Product list during ImportCustomObjects.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCombinedImportDemo
{
    // Sample custom object
    // Shows how to use Aspose.Cells for .NET to import a string array of category headers horizontally and a List&lt;Product&gt; of custom objects vertically into the same worksheet, display property names, apply a custom date format, insert rows, convert strings to numbers, auto‑fit columns, and save the file as an XLSX document.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Stock { get; set; }
        public DateTime ReleaseDate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // 1. Import a simple string array (e.g., category headers)
            // ------------------------------------------------------------
            string[] categories = new string[] { "Electronics", "Furniture", "Clothing", "Books" };
            // Import the array horizontally starting at cell B2 (row index 1, column index 1)
            // isVertical = false means horizontal placement
            cells.ImportArray(categories, 1, 1, false);

            // ------------------------------------------------------------
            // 2. Prepare a list of custom objects to import
            // ------------------------------------------------------------
            List<Product> products = new List<Product>
            {
                new Product { Name = "Laptop", Price = 999.99m, Stock = 25, ReleaseDate = new DateTime(2023, 5, 10) },
                new Product { Name = "Desk", Price = 199.50m, Stock = 40, ReleaseDate = new DateTime(2022, 11, 15) },
                new Product { Name = "T-Shirt", Price = 19.99m, Stock = 150, ReleaseDate = new DateTime(2023, 3, 1) },
                new Product { Name = "Novel", Price = 12.75m, Stock = 80, ReleaseDate = new DateTime(2021, 9, 20) }
            };

            // Define the property names to import (order matters)
            string[] propertyNames = { "Name", "Price", "Stock", "ReleaseDate" };

            // Import the custom objects starting below the previously imported categories.
            // First row index = 3 (row 4 in Excel), first column index = 0 (column A)
            // Show property names in the first row of this block (isPropertyNameShown = true)
            // Insert rows if needed, use a specific date format, and convert strings to numbers where possible.
            int importedRows = cells.ImportCustomObjects(
                products,                // ICollection list
                propertyNames,           // string[] propertyNames
                true,                    // bool isPropertyNameShown
                3,                       // int firstRow
                0,                       // int firstColumn
                products.Count,          // int rowNumber
                true,                    // bool insertRows
                "yyyy-MM-dd",            // string dateFormatString
                true                     // bool convertStringToNumber
            );

            // Optional: Auto-fit columns for better visibility
            worksheet.AutoFitColumns();

            // Save the workbook to an XLSX file
            workbook.Save("CombinedImportDemo.xlsx");
        }
    }
}
