// Title: Bind Multiple DataTables to Smart Marker Groups with WorkbookDesigner.SetDataSource (C# Aspose.Cells)
// Description: Demonstrates how to create a DataSet with several DataTables, place smart markers for each table in distinct worksheet sections, and use WorkbookDesigner.SetDataSource to automatically link each DataTable to its smart‑marker group. The example processes the markers and saves the result as SmartMarkerOutput.xlsx.
// Keywords: Aspose.Cells SetDataSource | C# smart markers multiple tables | WorkbookDesigner DataSet binding | smart marker groups Aspose.Cells | bind DataSet to smart markers .NET | Aspose.Cells multi‑table example
// Common Searches: How to bind a DataSet with several DataTables to different smart marker groups in Aspose.Cells | WorkbookDesigner SetDataSource example with multiple tables | Aspose.Cells smart markers using a DataSet | C# bind DataSet to smart markers for Customers and Orders
// Developer Intent: The developer needs to connect a DataSet that contains multiple DataTables to corresponding smart‑marker blocks in an Excel template using WorkbookDesigner.SetDataSource.
// Use Cases: Create a single worksheet that lists customers in one area and their orders in another, each populated from separate DataTables. | Generate a master‑detail report (e.g., employees and salaries) where each detail section is driven by its own DataTable. | Export related data sets such as Products, Categories, and Suppliers into one template, assigning each table to a distinct smart‑marker region.
// AI Prompts: Add a third DataTable for Products and place its smart markers below the Orders section in the same worksheet. | Show how to apply formatting (bold headers, auto‑fit columns) after designer.Process() completes. | Explain how to use SetDataSource with a DataSet that defines relationships and reference parent/child fields in smart markers.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Demonstrates how to create a DataSet with several DataTables, place smart markers for each table in distinct worksheet sections, and use WorkbookDesigner.SetDataSource to automatically link each DataTable to its smart‑marker group. The example processes the markers and saves the result as SmartMarkerOutput.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a DataSet containing multiple tables.
            DataSet ds = new DataSet();

            // First table: Customers
            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("CustomerID", typeof(int));
            customers.Columns.Add("Name", typeof(string));
            customers.Rows.Add(1, "John Doe");
            customers.Rows.Add(2, "Jane Smith");
            ds.Tables.Add(customers);

            // Second table: Orders
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("CustomerID", typeof(int));
            orders.Columns.Add("Product", typeof(string));
            orders.Columns.Add("Quantity", typeof(int));
            orders.Rows.Add(1001, 1, "Laptop", 2);
            orders.Rows.Add(1002, 2, "Smartphone", 5);
            ds.Tables.Add(orders);

            // Create a workbook and place smart markers for each table.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Smart markers for Customers table.
            ws.Cells["A1"].PutValue("CustomerID");
            ws.Cells["B1"].PutValue("Name");
            ws.Cells["A2"].PutValue("&=Customers.CustomerID");
            ws.Cells["B2"].PutValue("&=Customers.Name");

            // Smart markers for Orders table (starting at row 6).
            ws.Cells["A6"].PutValue("OrderID");
            ws.Cells["B6"].PutValue("CustomerID");
            ws.Cells["C6"].PutValue("Product");
            ws.Cells["D6"].PutValue("Quantity");
            ws.Cells["A7"].PutValue("&=Orders.OrderID");
            ws.Cells["B7"].PutValue("&=Orders.CustomerID");
            ws.Cells["C7"].PutValue("&=Orders.Product");
            ws.Cells["D7"].PutValue("&=Orders.Quantity");

            // Initialize the designer with the workbook.
            WorkbookDesigner designer = new WorkbookDesigner(wb);

            // Bind the DataSet; each DataTable is automatically linked to its smart marker group.
            designer.SetDataSource(ds);

            // Process the smart markers.
            designer.Process();

            // Save the populated workbook.
            designer.Workbook.Save("SmartMarkerOutput.xlsx");
        }
    }
}
