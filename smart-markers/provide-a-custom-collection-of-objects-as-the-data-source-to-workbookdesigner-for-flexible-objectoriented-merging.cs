// Title: C# – Bind a List<Product> to WorkbookDesigner Smart Markers with Aspose.Cells
// Description: Demonstrates how to create a workbook template, define smart markers that reference a "Products" collection, populate a List<Product> (Name, Price, Quantity), bind the list to WorkbookDesigner via SetDataSource, process the markers, and save the result as CustomObjectDataSource.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells WorkbookDesigner | C# smart markers custom collection | SetDataSource List<T> | bind custom objects to Excel template | Aspose.Cells data source example | Excel generation from object list | custom data source for smart markers | .NET Excel export
// Common Searches: How to bind a List of custom objects to WorkbookDesigner smart markers | Aspose.Cells SetDataSource generic List example | C# smart markers with custom class as data source | WorkbookDesigner process custom collection tutorial | Aspose.Cells custom object data source for Excel
// Developer Intent: Bind a custom List<Product> to WorkbookDesigner smart markers and generate a populated Excel file.
// Use Cases: Create an inventory sheet where each product row is filled from a List<Product>. | Generate an invoice that lists line‑item details using a custom collection. | Export sales or stock data to Excel by mapping DTO objects to smart markers.
// AI Prompts: Add a total row that calculates Σ(Price × Quantity) after processing the smart markers. | Show how to use a DataTable instead of List<Product> as the data source for the same markers. | Explain how to apply currency formatting to the Price column post‑processing.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomDataSource
{
    // Custom object that will be used as a data source
    // Demonstrates how to create a workbook template, define smart markers that reference a "Products" collection, populate a List<Product> (Name, Price, Quantity), bind the list to WorkbookDesigner via SetDataSource, process the markers, and save the result as CustomObjectDataSource.xlsx using Aspose.Cells for .NET.
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public int Quantity { get; set; }

        public Product(string name, double price, int quantity)
        {
            Name = name;
            Price = price;
            Quantity = quantity;
        }
    }

    public class Demo
    {
        public static void Main()
        {
            // ---------- Create a new workbook (template) ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Define smart markers for the collection ----------
            // Header row
            sheet.Cells["A1"].PutValue("Product Name");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["C1"].PutValue("Quantity");

            // Data rows – the markers reference the collection name "Products"
            sheet.Cells["A2"].PutValue("&=Products.Name");
            sheet.Cells["B2"].PutValue("&=Products.Price");
            sheet.Cells["C2"].PutValue("&=Products.Quantity");

            // ---------- Prepare a custom collection ----------
            List<Product> products = new List<Product>
            {
                new Product("Apple", 1.20, 50),
                new Product("Banana", 0.80, 100),
                new Product("Orange", 1.00, 75)
            };

            // ---------- Initialize WorkbookDesigner and assign the workbook ----------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            // ---------- Bind the collection to the smart marker variable ----------
            // This uses the SetDataSource(string, object) overload.
            designer.SetDataSource("Products", products);

            // Alternative approach (commented out):
            // Use CellsDataTableFactory to create an ICellsDataTable from the collection
            // CellsDataTableFactory factory = workbook.CellsDataTableFactory;
            // ICellsDataTable dataTable = factory.GetInstance(products);
            // designer.SetDataSource("Products", dataTable);

            // ---------- Process the smart markers ----------
            designer.Process();

            // ---------- Save the populated workbook ----------
            designer.Workbook.Save("CustomObjectDataSource.xlsx");
        }
    }
}
