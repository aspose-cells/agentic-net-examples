// Title: Aspose.Cells C# – Preserve Cell Formatting When Filling Smart Markers from a List
// Description: Demonstrates how to create a workbook template, apply custom styles, define range smart markers, bind a List<Product> to the marker name, and process the markers with WorkbookDesigner so that the original cell formatting (font color, currency, date format) is automatically retained. No LineByLine setting is required.
// Keywords: Aspose.Cells | C# | Smart markers | range smart markers | preserve cell formatting | WorkbookDesigner | custom object list | List<Product> | template styling | Excel export | retain style
// Common Searches: Aspose.Cells keep cell style when using smart markers | C# smart markers preserve formatting | range smart markers retain original formatting | WorkbookDesigner populate list without losing styles | how to avoid LineByLine for formatting in Aspose.Cells
// Developer Intent: Keep the predefined cell styles intact while populating smart markers from a collection of custom objects.
// Use Cases: Generate a product catalog where each row inherits the template's font colors and number formats. | Create a financial report that maintains currency and date formats after expanding rows from a List<Product>. | Build a styled inventory sheet using a named‑range smart marker so that added rows automatically match the original styling.
// AI Prompts: Show C# code that uses Aspose.Cells WorkbookDesigner to preserve cell formatting when processing smart markers from a List<T>. | Explain why the LineByLine option is unnecessary for formatting preservation with range smart markers in Aspose.Cells. | Provide an example of a named‑range smart marker that keeps original styles while binding a custom object collection.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerPreserveFormattingDemo
{
    // Sample custom object
    // Demonstrates how to create a workbook template, apply custom styles, define range smart markers, bind a List<Product> to the marker name, and process the markers with WorkbookDesigner so that the original cell formatting (font color, currency, date format) is automatically retained. No LineByLine setting is required.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public decimal Price { get; set; }
        public DateTime ReleaseDate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook (template)
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                Cells cells = ws.Cells;

                // 2. Set up smart markers in a named range
                // Header row
                cells["A1"].PutValue("Product Name");
                cells["B1"].PutValue("Price");
                cells["C1"].PutValue("Release Date");

                // Data row with smart markers
                cells["A2"].PutValue("&=Products.Name");
                cells["B2"].PutValue("&=Products.Price");
                cells["C2"].PutValue("&=Products.ReleaseDate");

                // Define the range that contains smart markers and give it the required name
                AsposeRange smartRange = cells.CreateRange("A2:C2");
                smartRange.Name = "_CellsSmartMarkers";

                // 3. Apply formatting to the template cells (these should be preserved)
                Style nameStyle = wb.CreateStyle();
                nameStyle.Font.Color = System.Drawing.Color.Blue;
                cells["A2"].SetStyle(nameStyle);

                Style priceStyle = wb.CreateStyle();
                priceStyle.Number = 2; // Currency format
                priceStyle.Font.Color = System.Drawing.Color.Green;
                cells["B2"].SetStyle(priceStyle);

                Style dateStyle = wb.CreateStyle();
                dateStyle.Custom = "dd-mmm-yyyy";
                dateStyle.Font.Color = System.Drawing.Color.Purple;
                cells["C2"].SetStyle(dateStyle);

                // 4. Prepare a list of custom objects
                List<Product> productList = new List<Product>
                {
                    new Product { Name = "Laptop", Price = 1299.99m, ReleaseDate = new DateTime(2023, 5, 10) },
                    new Product { Name = "Smartphone", Price = 799.50m, ReleaseDate = new DateTime(2023, 8, 22) },
                    new Product { Name = "Tablet", Price = 450.00m, ReleaseDate = new DateTime(2023, 11, 5) }
                };

                // 5. Configure WorkbookDesigner
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = wb
                    // No need to set LineByLine; using range smart markers preserves formatting.
                };

                // Bind the custom object list to the smart marker name "Products"
                designer.SetDataSource("Products", productList);

                // 6. Process the smart markers
                designer.Process();

                // 7. Save the result
                string outputPath = "SmartMarkerPreserveFormatting_Output.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
