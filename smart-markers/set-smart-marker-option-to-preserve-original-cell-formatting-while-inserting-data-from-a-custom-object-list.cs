using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public string Category { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook (template)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Define header cells with desired formatting -----
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["C1"].PutValue("Category");

            // Header style (bold, white font on dark blue background)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = System.Drawing.Color.White;
            headerStyle.ForegroundColor = System.Drawing.Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            sheet.Cells["A1"].SetStyle(headerStyle);
            sheet.Cells["B1"].SetStyle(headerStyle);
            sheet.Cells["C1"].SetStyle(headerStyle);

            // Number format for the price column (currency)
            Style priceStyle = workbook.CreateStyle();
            priceStyle.Number = 10; // Currency format
            sheet.Cells["B1"].SetStyle(priceStyle);

            // ----- Insert smart markers for data rows -----
            sheet.Cells["A2"].PutValue("&=Products.Name");
            sheet.Cells["B2"].PutValue("&=Products.Price");
            sheet.Cells["C2"].PutValue("&=Products.Category");

            // Define the range that contains the smart markers and give it a name
            Aspose.Cells.Range smartMarkerRange = sheet.Cells.CreateRange("A2:C2");
            smartMarkerRange.Name = "_CellsSmartMarkers";

            // ----- Prepare custom object list -----
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.23m, Category = "Fruit" },
                new Product { Name = "Laptop", Price = 999.99m, Category = "Electronics" },
                new Product { Name = "Book",   Price = 15.50m, Category = "Stationery" }
            };

            // ----- Set up WorkbookDesigner -----
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
                // No need to set LineByLine; using range smart markers preserves formatting
            };

            // Bind the custom object list to the smart marker name "Products"
            designer.SetDataSource("Products", products);

            // Process only the defined smart‑marker range.
            // The second parameter (true) indicates that unrecognized markers are not preserved.
            designer.Process(smartMarkerRange, true);

            // ----- Save the result -----
            string outputPath = "SmartMarkerPreserveFormatting.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}