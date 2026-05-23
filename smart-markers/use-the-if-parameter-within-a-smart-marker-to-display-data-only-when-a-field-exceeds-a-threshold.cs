using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerIfDemo
{
    // Sample data class
    public class Product
    {
        public string Name { get; set; }
        public double Sales { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook (template)
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // 2. Add header row
            ws.Cells["A1"].PutValue("Product");
            ws.Cells["B1"].PutValue("Sales (shown only if > 1000)");

            // 3. Insert smart markers.
            //    &IF($Sales>1000,$Sales,"") will display the Sales value only when it exceeds 1000.
            ws.Cells["A2"].PutValue("&=Products.Name");
            ws.Cells["B2"].PutValue("&=IF($Sales>1000,$Sales,\"\")");

            // 4. Define the range that contains smart markers and name it as required.
            //    This is needed when LineByLine is set to false (optional here).
            ws.Cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";

            // 5. Prepare sample data
            List<Product> products = new List<Product>
            {
                new Product { Name = "Widget", Sales = 850.0 },
                new Product { Name = "Gadget", Sales = 1250.5 },
                new Product { Name = "Doohickey", Sales = 3000.0 }
            };

            // 6. Set up the WorkbookDesigner and bind the data source.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = wb,
                LineByLine = false   // Use range smart markers
            };
            designer.SetDataSource("Products", products);

            // 7. Process the smart markers.
            designer.Process();

            // 8. Save the result.
            wb.Save("SmartMarkerIfDemo.xlsx");
        }
    }
}