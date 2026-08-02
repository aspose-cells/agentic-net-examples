// Title: C# – Fill Excel Smart Markers from a DataTable using WorkbookDesigner.SetDataSource (Aspose.Cells)
// Description: Demonstrates how to create an Excel template, add smart markers (e.g., "&=Products.ProductID"), build a DataTable that mimics a relational query result, bind the table to the workbook with WorkbookDesigner.SetDataSource, process the markers to expand rows and insert data, and save the populated file.
// Keywords: Aspose.Cells | WorkbookDesigner | SetDataSource | DataTable | smart markers | C# | .NET Excel export | template binding | relational data | Excel automation
// Common Searches: Aspose.Cells bind DataTable to smart markers | WorkbookDesigner.SetDataSource example C# | populate Excel template from database using Aspose.Cells | smart markers with DataTable in .NET | how to process smart markers in Aspose.Cells
// Developer Intent: The developer needs to populate an Excel template’s smart markers with rows from a DataTable using WorkbookDesigner.SetDataSource and then generate the final workbook.
// Use Cases: Generate a product catalog by pulling product rows from a database into a smart‑marked Excel sheet. | Create invoices where each line‑item is filled from an order‑details DataTable. | Export query results to a formatted sales report with dynamic row expansion via smart markers.
// AI Prompts: Show how to load an existing Excel template file instead of creating a new workbook. | Provide a sample that uses multiple related DataTables (e.g., Orders and OrderDetails) with hierarchical smart markers. | Explain how to apply currency formatting to the UnitPrice column while processing smart markers.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Demonstrates how to create an Excel template, add smart markers (e.g., "&=Products.ProductID"), build a DataTable that mimics a relational query result, bind the table to the workbook with WorkbookDesigner.SetDataSource, process the markers to expand rows and insert data, and save the populated file.
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a workbook that will act as the template.
            //    In a real scenario you would load an existing template file:
            //    Workbook workbook = new Workbook("Template.xlsx");
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();

            // -----------------------------------------------------------------
            // 2. Add smart markers to the worksheet.
            //    Smart markers are placed in cells using the syntax "&=DataSource.Column".
            // -----------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            // Header row
            sheet.Cells["A1"].PutValue("Product ID");
            sheet.Cells["B1"].PutValue("Product Name");
            sheet.Cells["C1"].PutValue("Unit Price");
            // Smart marker row – the designer will repeat this row for each data row.
            sheet.Cells["A2"].PutValue("&=Products.ProductID");
            sheet.Cells["B2"].PutValue("&=Products.ProductName");
            sheet.Cells["C2"].PutValue("&=Products.UnitPrice");

            // -----------------------------------------------------------------
            // 3. Simulate retrieving data from a relational database.
            //    Here we create a DataTable with the same schema as the expected result.
            // -----------------------------------------------------------------
            DataTable productTable = new DataTable("Products");
            productTable.Columns.Add("ProductID", typeof(int));
            productTable.Columns.Add("ProductName", typeof(string));
            productTable.Columns.Add("UnitPrice", typeof(decimal));

            // Sample rows – replace this block with actual DB query results.
            productTable.Rows.Add(101, "Chai", 18.00m);
            productTable.Rows.Add(102, "Chang", 19.00m);
            productTable.Rows.Add(103, "Aniseed Syrup", 10.00m);

            // -----------------------------------------------------------------
            // 4. Create a WorkbookDesigner, assign the workbook and set the DataTable
            //    as the data source. The SetDataSource(DataTable) overload binds the
            //    table name ("Products") automatically.
            // -----------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;
            designer.SetDataSource(productTable);

            // -----------------------------------------------------------------
            // 5. Process the smart markers – this expands the template rows and fills
            //    the cells with data from the DataTable.
            // -----------------------------------------------------------------
            designer.Process();

            // -----------------------------------------------------------------
            // 6. Save the populated workbook.
            // -----------------------------------------------------------------
            workbook.Save("SmartMarker_Output.xlsx");
        }
    }
}
