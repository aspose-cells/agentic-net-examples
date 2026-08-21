// Title: C# – Populate Excel Smart Markers from a DataTable with WorkbookDesigner.SetDataSource (Aspose.Cells)
// Description: Demonstrates how to create an Excel template, add smart markers like &=Products.ProductID, bind a DataTable named "Products" using WorkbookDesigner.SetDataSource, process the markers, and save the populated workbook as SmartMarkers_Output.xlsx.
// Keywords: Aspose.Cells | WorkbookDesigner | SetDataSource | DataTable | smart markers | C# Excel export | Excel template binding | relational data to Excel | Aspose.Cells example
// Common Searches: Aspose.Cells bind DataTable to smart markers | WorkbookDesigner SetDataSource C# example | populate Excel smart markers from database results | smart markers with DataTable Aspose | generate Excel report from relational data Aspose.Cells
// Developer Intent: Bind a DataTable to smart markers in an Excel template and generate a populated workbook using Aspose.Cells.
// Use Cases: Create a product catalog by mapping query results to smart markers. | Export invoice line items from a DataTable into a formatted spreadsheet. | Generate any relational query result as a styled Excel report with minimal code.
// AI Prompts: Provide C# code that uses WorkbookDesigner.SetDataSource with multiple related DataTables and processes smart markers. | Explain how the column names in a DataTable must match the &=TableName.ColumnName syntax for Aspose.Cells smart markers. | List troubleshooting steps when data does not appear after calling SetDataSource and Process.

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerDemo
{
    // Demonstrates how to create an Excel template, add smart markers like &=Products.ProductID, bind a DataTable named "Products" using WorkbookDesigner.SetDataSource, process the markers, and save the populated workbook as SmartMarkers_Output.xlsx.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook that will serve as the template.
            Workbook workbook = new Workbook();

            // 2. Add smart markers to the first worksheet.
            //    The markers follow the pattern &=TableName.ColumnName.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("&=Products.ProductID");   // Header marker (optional)
            sheet.Cells["B1"].PutValue("&=Products.ProductName"); // Header marker (optional)
            sheet.Cells["C1"].PutValue("&=Products.Price");       // Header marker (optional)

            // 3. Prepare a DataTable that simulates data retrieved from a relational database.
            DataTable productTable = new DataTable("Products");
            productTable.Columns.Add("ProductID", typeof(int));
            productTable.Columns.Add("ProductName", typeof(string));
            productTable.Columns.Add("Price", typeof(decimal));

            // Sample rows – in a real scenario these would come from a DB query.
            productTable.Rows.Add(101, "Laptop", 1200.50m);
            productTable.Rows.Add(102, "Smartphone", 799.99m);
            productTable.Rows.Add(103, "Tablet", 450.00m);

            // 4. Create a WorkbookDesigner and bind the workbook.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // 5. Set the DataTable as the data source for the smart markers.
            //    The table name ("Products") matches the marker prefix used above.
            designer.SetDataSource(productTable);

            // 6. Process the smart markers – this populates the worksheet with the data.
            designer.Process();

            // 7. Save the resulting workbook.
            workbook.Save("SmartMarkers_Output.xlsx");
        }
    }
}
