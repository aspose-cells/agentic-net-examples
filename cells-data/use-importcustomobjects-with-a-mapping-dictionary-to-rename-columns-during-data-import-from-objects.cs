// Title: Rename Excel column headers after ImportCustomObjects with Aspose.Cells (C#)
// Description: Demonstrates how to import a List<T> into an Excel worksheet using Cells.ImportCustomObjects, display the default property names, then replace those headers with custom titles from a dictionary mapping, apply a date format, and save the workbook.
// Keywords: Aspose.Cells ImportCustomObjects | C# rename Excel headers | dictionary column mapping | custom Excel column titles | export objects to Excel | date formatting Aspose.Cells
// Common Searches: Aspose.Cells change header names after ImportCustomObjects | C# map object properties to custom Excel column titles | how to rename columns when exporting a list to Excel | use dictionary to rename Excel headers Aspose.Cells
// Developer Intent: Import a collection of custom objects into a worksheet and substitute the automatically generated property‑name headers with user‑defined column titles.
// Use Cases: Create a product catalog where column headings need friendly names like "Product Name" and "Unit Price". | Export sales or inventory data to Excel and align column titles with a predefined reporting template. | Generate data files for external systems that require specific header names and date formats.
// AI Prompts: Show C# code that imports a list of objects with Cells.ImportCustomObjects and then renames the header row using a dictionary of custom column names. | Explain how to locate the header cells produced by ImportCustomObjects and replace their values with mapped titles in Aspose.Cells for .NET. | Provide an example that imports objects, formats date columns, and applies a header‑renaming dictionary before saving the workbook.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace ImportCustomObjectsWithRename
{
    // Sample data class
    // Demonstrates how to import a List<T> into an Excel worksheet using Cells.ImportCustomObjects, display the default property names, then replace those headers with custom titles from a dictionary mapping, apply a date format, and save the workbook.
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Stock { get; set; }
        public DateTime Date { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Prepare sample data
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 2.99m, Stock = 150, Date = new DateTime(2023,12,31) },
                new Product { Name = "Orange", Price = 1.99m, Stock = 200, Date = new DateTime(2024, 1, 15) }
            };

            // Mapping dictionary: original property name -> desired column header
            Dictionary<string, string> columnRenameMap = new Dictionary<string, string>
            {
                { "Name",  "Product Name" },
                { "Price", "Unit Price" },
                { "Stock", "Quantity In Stock" },
                { "Date",  "Release Date" }
            };

            // Import the custom objects.
            // propertyNames = null => import all properties.
            // isPropertyNameShown = true => first row will contain property names.
            // firstRow = 0, firstColumn = 0 => start at cell A1.
            // rowNumber = products.Count, insertRows = true, date format, convert strings to numbers.
            int importedRows = cells.ImportCustomObjects(
                products,
                null,
                true,
                0,
                0,
                products.Count,
                true,
                "yyyy-MM-dd",
                true);

            // Rename the header cells according to the mapping dictionary
            // Header row is at index 0 (firstRow)
            foreach (KeyValuePair<string, string> kvp in columnRenameMap)
            {
                // Find the column index of the original property name
                // Since we imported with property names shown, the header cell text matches the property name.
                // Scan the first row to locate the cell.
                for (int col = 0; col < cells.MaxColumn; col++)
                {
                    Cell headerCell = cells[0, col];
                    if (headerCell.Type == CellValueType.IsString && headerCell.StringValue == kvp.Key)
                    {
                        headerCell.PutValue(kvp.Value); // Replace with new name
                        break;
                    }
                }
            }

            // Save the workbook
            workbook.Save("ProductsRenamed.xlsx");
        }
    }
}
