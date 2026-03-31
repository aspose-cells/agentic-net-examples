using System;
using System.Collections.Generic;
using Aspose.Cells;

class SmartMarkerIfDemo
{
    static void Main()
    {
        // Create a new workbook that will serve as the template.
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Add smart markers to the template.
        //   A2: &Products.Name
        //   B2: &Products.Price
        //   C2: &IF($IsPremium, "Premium", "Standard")
        //   D2: &=$ReportDate
        ws.Cells["A2"].PutValue("&Products.Name");
        ws.Cells["B2"].PutValue("&Products.Price");
        ws.Cells["C2"].PutValue("&IF($IsPremium, \"Premium\", \"Standard\")");
        ws.Cells["D2"].PutValue("&=$ReportDate");

        // Create a WorkbookDesigner and associate it with the template workbook.
        WorkbookDesigner designer = new WorkbookDesigner(wb);

        // Prepare a collection that will be bound to the "Products" smart marker.
        List<Product> products = new List<Product>
        {
            new Product { Name = "Laptop",   Price = 1200.00, IsPremium = true  },
            new Product { Name = "Mouse",    Price =  25.00, IsPremium = false },
            new Product { Name = "Keyboard", Price =  45.00, IsPremium = false }
        };

        // Bind the collection to the smart marker name "Products".
        designer.SetDataSource("Products", products);

        // Set a variable that can be referenced directly in smart markers.
        // Variable name "ReportDate" will be used as &$ReportDate in the template.
        designer.SetDataSource("ReportDate", DateTime.Today);

        // Process all smart markers in the workbook.
        designer.Process();

        // Save the populated workbook.
        wb.Save("ResultWithIf.xlsx");
    }

    // Simple POCO class used as a data source for the "Products" smart marker.
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public bool IsPremium { get; set; }
    }
}