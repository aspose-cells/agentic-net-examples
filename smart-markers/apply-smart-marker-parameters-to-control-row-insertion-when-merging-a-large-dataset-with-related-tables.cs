// Title: C# Example: Controlling Row Insertion with Smart Marker ‘noadd’ for Parent‑Child DataTables in Aspose.Cells
// Description: This sample builds a DataSet with Customers (parent) and Orders (child) tables, defines a worksheet template that uses the smart‑marker parameter **noadd** on the order fields to stop automatic row creation, processes the markers to fill customer rows, then programmatically inserts the exact number of order rows for each customer and imports the order data into columns C and D. The workbook is saved as an Excel file.
// Keywords: Aspose.Cells smart markers noadd | C# parent child DataSet Excel | manual row insertion Aspose.Cells | ImportData ImportTableOptions | .NET Excel template smart markers | GitHub Aspose.Cells example | Excel report customer orders
// Common Searches: Aspose.Cells prevent automatic row insertion for child table | how to use noadd smart marker in C# | insert rows manually after processing smart markers | merge related DataTables into Excel with Aspose.Cells | sample code for parent‑detail smart markers
// Developer Intent: Generate an Excel report from a parent‑child DataSet while suppressing the default row expansion for the child table and adding rows only where needed using the smart‑marker **noadd** parameter and programmatic row insertion.
// Use Cases: Customer‑order summary where each customer appears once and their orders are listed directly beneath, with rows added only for existing orders. | Invoice generation that adds line‑item rows dynamically after the main invoice data has been populated. | Event schedule that groups events under categories, inserting rows only for categories that contain events.
// AI Prompts: Write C# code using Aspose.Cells to apply the ‘noadd’ smart marker to a child table, then insert the required rows per parent record and import the child data. | Explain the effect of the ‘noadd’ parameter on smart‑marker processing and demonstrate how to use ImportData with ImportTableOptions for manual row insertion. | Provide a step‑by‑step guide to create an Excel template with smart markers for a master‑detail relationship and control row insertion for the detail rows in .NET.

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerRowInsertionDemo
{
    // This sample builds a DataSet with Customers (parent) and Orders (child) tables, defines a worksheet template that uses the smart‑marker parameter **noadd** on the order fields to stop automatic row creation, processes the markers to fill customer rows, then programmatically inserts the exact number of order rows for each customer and imports the order data into columns C and D. The workbook is saved as an Excel file.
    class Program
    {
        static void Main()
        {
            // ---------- 1. Prepare a DataSet with related tables ----------
            DataSet ds = new DataSet();

            // Parent table: Customers
            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("CustomerID", typeof(int));
            customers.Columns.Add("CustomerName", typeof(string));
            customers.Rows.Add(1, "Alpha Corp");
            customers.Rows.Add(2, "Beta Ltd");
            ds.Tables.Add(customers);

            // Child table: Orders (related to Customers via CustomerID)
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("CustomerID", typeof(int));
            orders.Columns.Add("OrderDate", typeof(DateTime));
            orders.Rows.Add(1001, 1, new DateTime(2023, 1, 15));
            orders.Rows.Add(1002, 1, new DateTime(2023, 2, 20));
            orders.Rows.Add(2001, 2, new DateTime(2023, 3, 5));
            ds.Tables.Add(orders);

            // ---------- 2. Create a workbook template with smart markers ----------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Header row
            cells["A1"].PutValue("Customer ID");
            cells["B1"].PutValue("Customer Name");
            cells["C1"].PutValue("Order ID");
            cells["D1"].PutValue("Order Date");

            // Row for customer data (will be repeated for each customer)
            cells["A2"].PutValue("&=$Customers.CustomerID");
            cells["B2"].PutValue("&=$Customers.CustomerName");

            // Row for order data (smart marker with 'noadd' to suppress automatic row insertion)
            // The 'noadd' parameter tells the designer not to insert rows for each order record.
            cells["C2"].PutValue("&=$Orders.OrderID(noadd)");
            cells["D2"].PutValue("&=$Orders.OrderDate(noadd)");

            // ---------- 3. Process smart markers ----------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = wb
            };
            designer.SetDataSource(ds);
            designer.Process(); // Fills customer rows; order cells remain empty because of 'noadd'

            // ---------- 4. Manually insert rows for orders and import order data ----------
            // Determine where the first order row should be placed (row index 1 = second row, zero‑based)
            int orderStartRow = 1; // corresponds to Excel row 2 (the row with order smart markers)

            // For each customer, find related orders and insert rows accordingly
            foreach (DataRow custRow in customers.Rows)
            {
                int custId = (int)custRow["CustomerID"];

                // Filter orders for the current customer
                DataRow[] custOrders = orders.Select($"CustomerID = {custId}");

                if (custOrders.Length == 0)
                {
                    continue; // No orders for this customer
                }

                // Insert required number of rows below the current order placeholder row
                // (InsertRows inserts *before* the specified index, so we add after the placeholder)
                cells.InsertRows(orderStartRow + 1, custOrders.Length - 1, true);
                // The placeholder row already exists; we need only (count‑1) additional rows.

                // Prepare a temporary DataTable to hold the orders for this customer
                DataTable tempOrders = new DataTable();
                tempOrders.Columns.Add("OrderID", typeof(int));
                tempOrders.Columns.Add("OrderDate", typeof(DateTime));
                foreach (DataRow o in custOrders)
                {
                    tempOrders.Rows.Add(o["OrderID"], o["OrderDate"]);
                }

                // Import the order data starting at the placeholder row
                ImportTableOptions importOpts = new ImportTableOptions
                {
                    InsertRows = true,          // Ensure rows are added if needed (safety)
                    IsFieldNameShown = false   // Do not import column names again
                };
                // Import only the two columns (OrderID, OrderDate) into columns C and D
                cells.ImportData(tempOrders, orderStartRow, 2, importOpts);

                // Move the start row pointer past the rows we just filled
                orderStartRow += custOrders.Length;
            }

            // ---------- 5. Save the result ----------
            wb.Save("SmartMarkerRowInsertionResult.xlsx");
        }
    }
}
