// Title: Use WorkbookDesigner to bind an IEnumerable<Product> collection to smart markers in an Aspose.Cells workbook (C#)
// Common Searches: C# Aspose.Cells WorkbookDesigner bind IEnumerable collection to smart markers | How to populate smart markers from a List of objects using Aspose.Cells | Set data source for smart markers with IEnumerable<Product> in Aspose.Cells .NET | Aspose.Cells example: fill smart markers from a collection across worksheets | Populate Excel smart markers using WorkbookDesigner and IEnumerable in C#
// Tags: WorkbookDesigner set data source IEnumerable | Aspose.Cells smart markers collection binding | populate smart markers from List<Product> | C# Excel smart markers using Aspose.Cells | multiple worksheets smart marker population

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerDemo
{
    // Sample data class
    // The example creates a new workbook, adds smart marker placeholders, prepares an IEnumerable<Product> list, configures WorkbookDesigner with the collection name matching the smart marker prefix, processes the workbook to replace the markers with actual data, and saves the populated Excel file.
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // 2. Add smart markers to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Products";

            // Header row
            sheet.Cells["A1"].PutValue("Product Name");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["C1"].PutValue("Quantity");

            // Smart marker rows (the = indicates a smart marker)
            sheet.Cells["A2"].PutValue("&=Products.Name");
            sheet.Cells["B2"].PutValue("&=Products.Price");
            sheet.Cells["C2"].PutValue("&=Products.Quantity");

            // 3. Prepare an IEnumerable collection as data source
            IEnumerable<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 0.5, Quantity = 120 },
                new Product { Name = "Banana", Price = 0.3, Quantity = 200 },
                new Product { Name = "Orange", Price = 0.7, Quantity = 150 }
            };

            // 4. Configure WorkbookDesigner to use the collection (feature rule: chart-values not applicable, so free‑form)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // The collection name "Products" must match the smart marker prefix
            designer.SetDataSource("Products", products);
            designer.Process(); // Populate smart markers with data

            // 5. Save the populated workbook (lifecycle rule: save)
            workbook.Save("SmartMarkers_Output.xlsx");
        }
    }
}
