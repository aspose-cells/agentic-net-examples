using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add headers
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");

        // Add smart markers.
        // Column A: simple field insertion.
        sheet.Cells["A2"].PutValue("&=Products.Name");
        // Column B: display Sales only when it exceeds 500 using IF smart marker.
        sheet.Cells["B2"].PutValue("&=IF($Sales>500,$Sales,\"\")");

        // Define the range that contains smart markers (required for processing).
        sheet.Cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";

        // Prepare sample data source.
        List<Product> products = new List<Product>
        {
            new Product { Name = "Alpha", Sales = 300 },
            new Product { Name = "Beta",  Sales = 750 },
            new Product { Name = "Gamma", Sales = 1200 }
        };

        // Set up WorkbookDesigner and assign the data source (process rule will use it).
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };
        designer.SetDataSource("Products", products);

        // Process all smart markers (process rule).
        designer.Process();

        // Save the result (save rule).
        workbook.Save("IfSmartMarkerOutput.xlsx");
    }
}

// Simple POCO class representing the data source.
public class Product
{
    public string Name { get; set; }
    public double Sales { get; set; }
}