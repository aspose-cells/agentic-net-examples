using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerCurrencyFormat
{
    // Simple data class for demonstration
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }   // Currency value
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook and set up smart markers
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");

            // Smart marker rows (will be populated by WorkbookDesigner)
            // &=$Products.Name and &=$Products.Price are smart markers
            cells["A2"].PutValue("&=$Products.Name");
            cells["B2"].PutValue("&=$Products.Price");

            // Define the range that contains the smart markers
            // This range will be processed and later formatted
            Aspose.Cells.Range smRange = cells.CreateRange("A2:B2");
            smRange.Name = "_CellsSmartMarkers";

            // 2. Prepare sample data source
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.25 },
                new Product { Name = "Banana", Price = 0.75 },
                new Product { Name = "Cherry", Price = 2.10 }
            };

            // 3. Set up WorkbookDesigner, assign data source and process smart markers
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = wb,
                LineByLine = false   // Process all rows at once
            };
            designer.SetDataSource("Products", products);
            designer.Process();   // Process all smart markers in the workbook

            // 4. Apply a custom currency number format to the populated price column
            // Define the custom format (e.g., US dollars with two decimals)
            string customCurrencyFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";

            // Create a style with the custom format
            Style currencyStyle = wb.CreateStyle();
            currencyStyle.SetCustom(customCurrencyFormat, true);

            // Use StyleFlag to apply only the number format part
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the entire price column where data was inserted
            // After processing, the data occupies rows 2..(products.Count+1)
            int lastRow = 1 + products.Count; // 1-based index for header row
            Aspose.Cells.Range priceRange = cells.CreateRange(1, 1, products.Count, 1); // start at B2 (row=1,col=1)
            priceRange.ApplyStyle(currencyStyle, flag);

            // 5. Save the resulting workbook
            wb.Save("SmartMarkerCurrencyFormatted.xlsx");
        }
    }
}