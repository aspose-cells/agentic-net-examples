// Title: How to implement a custom ICellsDataTable in C# to feed Aspose.Cells smart markers from a web‑service
// AI Prompts: Create a C# class that implements ICellsDataTable to expose a List<Product> for Aspose.Cells smart markers. | Demonstrate binding the custom ICellsDataTable to WorkbookDesigner, inserting smart markers in a worksheet, and generating an Excel file. | Extend the ICellsDataTable to add a new column (e.g., Category) and update the template markers accordingly.
// Common Searches: asp.net example for custom ICellsDataTable with Aspose.Cells smart markers | populate Excel template from web service using Aspose.Cells WorkbookDesigner C# | how to bind a list of objects as a data source for smart markers in Aspose.Cells | extend ICellsDataTable implementation to include additional columns in Aspose.Cells | using ICellsDataTable to map JSON response to Excel smart markers in C#
// Tags: ICellsDataTable implementation for smart markers | custom data source binding WorkbookDesigner C# | populate Excel from web service Aspose.Cells | add dynamic columns to ICellsDataTable | smart marker template generation Aspose.Cells

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomDataSourceDemo
{
    // Simple POCO representing data returned from a web service
    // Demonstrates a C# ICellsDataTable that wraps a List<Product> returned from a simulated web service, binds it to WorkbookDesigner, uses smart markers (&=$Products.Name, &=$Products.Price) in a worksheet, processes the template, and saves the populated workbook as output.xlsx.
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
    }

    // Custom data source implementing ICellsDataTable.
    // This class adapts a list of Product objects so that WorkbookDesigner can read values via smart markers.
    public class ProductDataSource : ICellsDataTable
    {
        private readonly List<Product> _products;
        private int _currentRow = -1; // Position for enumeration

        public ProductDataSource(List<Product> products)
        {
            _products = products ?? new List<Product>();
        }

        // Indexer by row and column (0‑based). Column 0 = Name, Column 1 = Price.
        public object this[int rowIndex, int columnIndex]
        {
            get
            {
                if (rowIndex < 0 || rowIndex >= RowCount)
                    throw new IndexOutOfRangeException();

                return columnIndex == 0
                    ? (object)_products[rowIndex].Name
                    : (object)_products[rowIndex].Price;
            }
        }

        // Indexer by row only – returns the whole object.
        public object this[int rowIndex] => _products[rowIndex];

        // Indexer by column name – used when smart marker refers to a column name.
        public object this[string columnName]
        {
            get
            {
                if (_currentRow < 0 || _currentRow >= RowCount)
                    throw new InvalidOperationException("Enumeration has not started or has finished.");

                return columnName switch
                {
                    "Name" => _products[_currentRow].Name,
                    "Price" => _products[_currentRow].Price,
                    _ => throw new ArgumentException($"Column '{columnName}' does not exist.")
                };
            }
        }

        // Total number of rows.
        public int RowCount => _products.Count;

        // Number of columns (Name, Price).
        public int ColumnCount => 2;

        // Alias for RowCount – required by ICellsDataTable.
        public int Count => RowCount;

        // Column names in the order they appear.
        public string[] Columns => new[] { "Name", "Price" };

        // Reset enumeration to before the first row.
        public void BeforeFirst()
        {
            _currentRow = -1;
        }

        // Move to the next row; returns false when the end is reached.
        public bool Next()
        {
            _currentRow++;
            return _currentRow < RowCount;
        }
    }

    class Program
    {
        // Simulated web service call – in a real scenario this would be an async HTTP request.
        static List<Product> GetProductsFromWebService()
        {
            // Mocked response
            return new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.20 },
                new Product { Name = "Banana", Price = 0.80 },
                new Product { Name = "Cherry", Price = 2.50 }
            };
        }

        static void Main()
        {
            // ---------- Create a workbook (template) ----------
            Workbook workbook = new Workbook();                     // create
            Worksheet sheet = workbook.Worksheets[0];

            // Place smart markers that will be populated from the custom data source.
            // &=$Products.Name  -> column "Name"
            // &=$Products.Price -> column "Price"
            sheet.Cells["A1"].PutValue("&=$Products.Name");
            sheet.Cells["B1"].PutValue("&=$Products.Price");

            // ---------- Obtain data from the (simulated) web service ----------
            List<Product> products = GetProductsFromWebService();

            // ---------- Set up WorkbookDesigner with the custom data source ----------
            WorkbookDesigner designer = new WorkbookDesigner();     // create
            designer.Workbook = workbook;                           // assign workbook

            // Bind the custom ICellsDataTable implementation to the smart marker name "Products".
            designer.SetDataSource("Products", new ProductDataSource(products));

            // Process smart markers – this will fill the worksheet with data from the web service.
            designer.Process();

            // ---------- Save the populated workbook ----------
            workbook.Save("output.xlsx");                           // save
        }
    }
}
