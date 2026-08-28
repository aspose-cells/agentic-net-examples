// Title: How to bind a multi‑table DataSet to smart markers using WorkbookDesigner.SetDataSource in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a DataSet with 'Customers' and 'Orders' tables, assigns it to WorkbookDesigner via SetDataSource, and processes smart markers to generate an Excel workbook. | Show the step‑by‑step process for mapping different smart marker groups to separate DataTables in a DataSet and producing the populated file with Aspose.Cells. | Provide a complete example that adds smart marker placeholders, builds a DataSet containing several tables, binds it using WorkbookDesigner.SetDataSource, and saves the resulting Excel document.
// Common Searches: asp.net aspose.cells setdatasource multiple datatables smart markers example | c# bind dataset with customers and orders tables to workbookdesigner smart markers | how to populate smart markers from two tables in Aspose.Cells | using WorkbookDesigner.SetDataSource with a DataSet that has several tables | smart markers for customers and orders excel generation Aspose.Cells C#
// Tags: WorkbookDesigner SetDataSource with DataSet | populate smart markers from multiple DataTables | Aspose.Cells generate Excel from DataSet | C# smart markers multiple tables example | Excel smart markers data binding Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

// The sample creates a workbook, inserts header cells and smart marker placeholders referencing "Customers" and "Orders" tables, builds a DataSet containing two DataTables with sample data, binds the DataSet to a WorkbookDesigner via SetDataSource, processes the smart markers to fill the worksheet, and saves the populated workbook as an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook that will hold the smart markers.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header cells.
        sheet.Cells["A1"].PutValue("Customer Name");
        sheet.Cells["B1"].PutValue("Customer City");
        sheet.Cells["C1"].PutValue("Order ID");
        sheet.Cells["D1"].PutValue("Product");

        // Add smart marker rows. Each marker refers to a table name in the DataSet.
        // "Customers" and "Orders" are the two tables we will bind later.
        sheet.Cells["A2"].PutValue("&=Customers.Name");
        sheet.Cells["B2"].PutValue("&=Customers.City");
        sheet.Cells["C2"].PutValue("&=Orders.OrderID");
        sheet.Cells["D2"].PutValue("&=Orders.Product");

        // Build a DataSet containing two DataTables.
        DataSet dataSet = new DataSet();

        // First table: Customers
        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("Name", typeof(string));
        customers.Columns.Add("City", typeof(string));
        customers.Rows.Add("Alice", "London");
        customers.Rows.Add("Bob", "Paris");
        dataSet.Tables.Add(customers);

        // Second table: Orders
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("OrderID", typeof(int));
        orders.Columns.Add("Product", typeof(string));
        orders.Rows.Add(1001, "Laptop");
        orders.Rows.Add(1002, "Smartphone");
        dataSet.Tables.Add(orders);

        // Bind the DataSet to the WorkbookDesigner.
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dataSet);
        designer.Process(); // Populate the smart markers with data.

        // Save the populated workbook.
        designer.Workbook.Save("SmartMarkerMultipleTables.xlsx");
    }
}
