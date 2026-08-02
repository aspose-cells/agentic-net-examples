// Title: C# – Bind a DataSet with multiple DataTables to smart marker groups using WorkbookDesigner.SetDataSource (Aspose.Cells)
// Description: This example creates a DataSet with two DataTables (Orders and Customers), defines smart markers for each table in a workbook, and uses WorkbookDesigner.SetDataSource to automatically populate both tables in a single Excel file. It shows how to group smart markers, process the workbook, and save the result.
// Keywords: Aspose.Cells | WorkbookDesigner | SetDataSource | DataSet | DataTable | smart markers | multiple tables | C# | .NET | Excel report | example | tutorial | bind dataset | smart marker groups | populate Excel
// Common Searches: Aspose.Cells bind DataSet to smart markers | WorkbookDesigner SetDataSource multiple tables C# | How to use smart markers with several DataTables | Populate Excel with orders and customers using Aspose.Cells | Smart marker groups example Aspose.Cells .NET | C# Aspose.Cells generate report from DataSet
// Developer Intent: I need to map several DataTables in a DataSet to different smart‑marker groups in an Excel workbook using Aspose.Cells.
// Use Cases: Create a combined order and customer report in one worksheet. | Generate separate sheets for each DataTable while using a single SetDataSource call. | Automate export of relational data (e.g., orders, customers, products) to a formatted Excel workbook. | Produce invoices or purchase orders where each section pulls data from its own table. | Build a master workbook for business‑intelligence dashboards directly from a DataSet.
// AI Prompts: Show how to apply a custom number format to the Quantity column in the Orders smart‑marker block. | Provide a version of the code that writes the workbook to a MemoryStream instead of saving to disk. | Explain how to handle DataSet tables whose column names differ from the smart‑marker field names. | Demonstrate placing each DataTable on a separate worksheet while still using a single SetDataSource call. | Add conditional formatting to highlight high‑quantity orders in the generated Excel file.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a DataSet with two DataTables (Orders and Customers), defines smart markers for each table in a workbook, and uses WorkbookDesigner.SetDataSource to automatically populate both tables in a single Excel file. It shows how to group smart markers, process the workbook, and save the result.
    public class SetDataSourceWithMultipleTablesDemo
    {
        public static void Run()
        {
            try
            {
                // ---------- Create a DataSet with multiple DataTables ----------
                DataSet dataSet = new DataSet();

                // First table: Orders
                DataTable ordersTable = new DataTable("Orders");
                ordersTable.Columns.Add("OrderID", typeof(int));
                ordersTable.Columns.Add("Product", typeof(string));
                ordersTable.Columns.Add("Quantity", typeof(int));

                ordersTable.Rows.Add(1001, "Laptop", 2);
                ordersTable.Rows.Add(1002, "Smartphone", 5);
                ordersTable.Rows.Add(1003, "Tablet", 3);

                // Second table: Customers
                DataTable customersTable = new DataTable("Customers");
                customersTable.Columns.Add("CustomerID", typeof(int));
                customersTable.Columns.Add("Name", typeof(string));
                customersTable.Columns.Add("Country", typeof(string));

                customersTable.Rows.Add(1, "John Doe", "USA");
                customersTable.Rows.Add(2, "Anna Smith", "UK");
                customersTable.Rows.Add(3, "Li Wei", "China");

                // Add tables to the DataSet
                dataSet.Tables.Add(ordersTable);
                dataSet.Tables.Add(customersTable);

                // ---------- Prepare a workbook with smart markers ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Smart markers for the Orders table (group name "Orders")
                sheet.Cells["A1"].PutValue("Order ID");
                sheet.Cells["B1"].PutValue("Product");
                sheet.Cells["C1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("&=Orders.OrderID");
                sheet.Cells["B2"].PutValue("&=Orders.Product");
                sheet.Cells["C2"].PutValue("&=Orders.Quantity");

                // Smart markers for the Customers table (group name "Customers")
                sheet.Cells["E1"].PutValue("Customer ID");
                sheet.Cells["F1"].PutValue("Name");
                sheet.Cells["G1"].PutValue("Country");
                sheet.Cells["E2"].PutValue("&=Customers.CustomerID");
                sheet.Cells["F2"].PutValue("&=Customers.Name");
                sheet.Cells["G2"].PutValue("&=Customers.Country");

                // ---------- Bind the DataSet to the workbook designer ----------
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSet); // Map DataSet tables to smart marker groups
                designer.Process(); // Populate data

                // ---------- Save the result ----------
                workbook.Save("MultipleTablesOutput.xlsx");
                Console.WriteLine("Workbook saved as MultipleTablesOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetDataSourceWithMultipleTablesDemo.Run();
        }
    }
}
