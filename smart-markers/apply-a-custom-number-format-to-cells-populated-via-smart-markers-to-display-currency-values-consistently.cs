// Title: How to apply a custom accounting currency number format to smart‑marker generated price cells using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a custom accounting style with a dollar sign and applies it only to the number format of a cell range populated by WorkbookDesigner smart markers. | Show how to calculate the dynamic range of the price column based on a List<Product> and apply a custom currency format using Style and StyleFlag in Aspose.Cells. | Provide a complete Aspose.Cells example that inserts smart markers, processes them, and formats the resulting price cells with a custom accounting number format before saving the workbook.
// Common Searches: Aspose.Cells C# apply custom accounting number format to smart marker output | How to set a custom currency format for cells populated by WorkbookDesigner | C# example of using StyleFlag to change only number format after processing smart markers | Apply custom dollar sign format to a dynamic range in Aspose.Cells workbook
// Tags: custom accounting number format Aspose.Cells | smart markers currency formatting C# | apply StyleFlag number format range Aspose.Cells | dynamic range formatting based on List<Product> | WorkbookDesigner custom number format example

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSmartMarkerCurrencyDemo
{
    // Simple data class for demonstration
    // The example creates a workbook, adds smart markers for product names and prices, populates them from a List<Product> using WorkbookDesigner, defines a custom accounting currency style, determines the price column range dynamically, applies the style with a StyleFlag that targets only the number format, and saves the workbook as an Excel file.
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }   // Currency value
    }

    public class Program
    {
        public static void Main()
        {
            // -------------------------------------------------
            // 1. Create a workbook and add smart markers
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];               // first worksheet

            // Header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");

            // Smart markers – they will be replaced by data source values
            sheet.Cells["A2"].PutValue("&=Products.Name");
            sheet.Cells["B2"].PutValue("&=Products.Price");

            // -------------------------------------------------
            // 2. Prepare data source (list of products)
            // -------------------------------------------------
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.25 },
                new Product { Name = "Banana", Price = 0.80 },
                new Product { Name = "Cherry", Price = 2.50 }
            };

            // -------------------------------------------------
            // 3. Set data source and process smart markers
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                LineByLine = false               // process all rows at once
            };
            designer.SetDataSource("Products", products);
            designer.Process();                  // populate cells with data

            // -------------------------------------------------
            // 4. Apply custom currency number format to the populated price cells
            // -------------------------------------------------
            // Define custom format (Accounting style with dollar sign)
            string customCurrencyFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";

            // Create a style and set the custom format
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.SetCustom(customCurrencyFormat, true);   // true => use builtin if matches

            // Prepare a StyleFlag to apply only the number format
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;          // apply only number format part

            // Determine the range that contains the price values (column B, rows 2..N)
            int startRow = 1;                   // zero‑based index for row 2
            int priceColumn = 1;                // column B
            int rowCount = products.Count;      // number of data rows

            // Create the range and apply the style
            Aspose.Cells.Range priceRange = sheet.Cells.CreateRange(startRow, priceColumn, rowCount, 1);
            priceRange.ApplyStyle(currencyStyle, flag);

            // -------------------------------------------------
            // 5. Save the result
            // -------------------------------------------------
            workbook.Save("SmartMarkerCurrencyOutput.xlsx");
        }
    }
}
