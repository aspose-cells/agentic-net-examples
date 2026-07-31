// Title: Aspose.Cells C# – Preserve Cell Formatting When Using Smart Markers with a Custom Object List
// Description: Demonstrates how to keep original cell styles (bold headers, currency and date formats) while populating smart markers from a List<Product>. The example creates a template workbook, applies a named range "_CellsSmartMarkers", sets WorkbookDesigner.LineByLine = false, binds the data source, processes the markers, and saves the result.
// Keywords: Aspose.Cells smart markers | preserve cell formatting | WorkbookDesigner LineByLine false | named range _CellsSmartMarkers | C# custom object list | retain number format | retain date format | Excel template styling | Aspose.Cells example
// Common Searches: keep cell style when using Aspose.Cells smart markers | Aspose.Cells preserve number format with smart markers | LineByLine false smart markers example | named range _CellsSmartMarkers Aspose.Cells | C# smart markers formatting issue
// Developer Intent: Insert data via smart markers without altering the predefined cell formatting.
// Use Cases: Generate a product catalog where header fonts stay bold and gray, and price/date columns retain currency and short‑date formats. | Create an invoice sheet that repeats a styled row for each line item while preserving amount and due‑date formats. | Export an employee directory from a List<Employee> while maintaining custom fonts and hire‑date styles.
// AI Prompts: Show how to configure WorkbookDesigner with LineByLine = false and a named range _CellsSmartMarkers to keep cell formatting when processing smart markers from a List<T>. | Provide a concise C# example that binds a List<Product> to a smart marker and preserves number/date styles after designer.Process(). | Explain why setting LineByLine to false retains original cell styles and how to apply the same technique to percentages or custom number formats.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerPreserveFormattingDemo
{
    // Sample custom object
    // Demonstrates how to keep original cell styles (bold headers, currency and date formats) while populating smart markers from a List<Product>. The example creates a template workbook, applies a named range "_CellsSmartMarkers", sets WorkbookDesigner.LineByLine = false, binds the data source, processes the markers, and saves the result.
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
                // 1. Create a workbook that will act as a template
                Workbook templateWb = new Workbook();
                Worksheet sheet = templateWb.Worksheets[0];

                // Header row with formatting that we want to keep
                sheet.Cells["A1"].PutValue("Product Name");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["C1"].PutValue("Release Date");

                // Apply header style
                Style headerStyle = templateWb.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
                headerStyle.Pattern = BackgroundType.Solid;
                sheet.Cells["A1:C1"].SetStyle(headerStyle);

                // Insert smart markers in the row that will be repeated
                sheet.Cells["A2"].PutValue("&=Products.Name");
                sheet.Cells["B2"].PutValue("&=Products.Price");
                sheet.Cells["C2"].PutValue("&=Products.ReleaseDate");

                // Apply number/date formatting that should be preserved after data insertion
                Style dataStyle = templateWb.CreateStyle();
                dataStyle.Number = 14; // Currency format
                sheet.Cells["B2"].SetStyle(dataStyle);

                Style dateStyle = templateWb.CreateStyle();
                dateStyle.Number = 14; // Short date format
                sheet.Cells["C2"].SetStyle(dateStyle);

                // Define the range that contains the smart markers and give it the required name
                // When LineByLine is false, the range must be named "_CellsSmartMarkers"
                Aspose.Cells.Range smartRange = sheet.Cells.CreateRange("A2:C2");
                smartRange.Name = "_CellsSmartMarkers";

                // 2. Prepare a list of custom objects
                List<Product> products = new List<Product>
                {
                    new Product { Name = "Laptop", Price = 1299.99m, ReleaseDate = new DateTime(2023, 5, 10) },
                    new Product { Name = "Smartphone", Price = 799.50m, ReleaseDate = new DateTime(2023, 8, 22) },
                    new Product { Name = "Tablet", Price = 450.00m, ReleaseDate = new DateTime(2023, 11, 5) }
                };

                // 3. Set up WorkbookDesigner
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = templateWb,
                    // Setting LineByLine to false tells the designer to use the named range.
                    // This mode preserves the original formatting of the cells.
                    LineByLine = false
                };

                // Bind the custom object list to the smart marker name "Products"
                designer.SetDataSource("Products", products);

                // 4. Process the smart markers
                designer.Process();

                // 5. Save the result
                string outputPath = "SmartMarkerPreserveFormattingResult.xlsx";
                designer.Workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
