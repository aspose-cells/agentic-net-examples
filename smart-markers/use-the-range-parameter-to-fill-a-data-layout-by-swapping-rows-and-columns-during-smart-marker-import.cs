using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerTransposeDemo
{
    // Sample data class
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                var cells = sheet.Cells;

                // 2. Insert smart markers in a vertical layout (rows)
                //    Row 1: Header
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Price");
                //    Row 2: Smart markers for data rows
                cells["A2"].PutValue("&=$Products.Name");
                cells["B2"].PutValue("&=$Products.Price");

                // 3. Define a range that contains the smart markers (including header)
                //    This range will be transposed so that data is filled horizontally instead of vertically
                Aspose.Cells.Range smartRange = cells.CreateRange("A1:B2");
                // Name the range as required for smart marker processing
                smartRange.Name = "_CellsSmartMarkers";

                // 4. Transpose the range – rows become columns and vice‑versa
                smartRange.Transpose();

                // 5. Prepare data source (a list of products)
                var products = new List<Product>
                {
                    new Product { Name = "Laptop", Price = 1200.50 },
                    new Product { Name = "Phone",  Price =  799.99 },
                    new Product { Name = "Tablet", Price =  450.75 }
                };

                // 6. Set up WorkbookDesigner and assign the data source
                var designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Products", products);

                // 7. Process only the defined range (true = preserve unrecognized markers)
                designer.Process(smartRange, true);

                // 8. Save the result
                const string outputPath = "SmartMarkerTransposeOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}