// Title: Create a master‑detail Excel report with hierarchical smart markers and a DataSet using Aspose.Cells in C#
// AI Prompts: Write C# code that defines an Excel template with {Customers.Name} and {Customers.City} smart markers for the master section and {Customers.Orders.OrderID}, {Customers.Orders.OrderDate}, {Customers.Orders.Amount} markers for the detail section, then populate it from a DataSet and save the workbook as .xlsx. | Apply a date format (e.g., "mm/dd/yyyy") to the OrderDate column and a currency format to the Amount column in the smart‑marker template before processing it with WorkbookDesigner. | Insert a blank row between the master and detail tables, auto‑size the columns after processing, and export the final file as an Excel workbook.
// Common Searches: asp.net c# generate master detail Excel using Aspose.Cells smart markers | how to use hierarchical smart markers with a DataSet in Aspose.Cells | Aspose.Cells WorkbookDesigner master‑detail report example | C# create Excel report with customers and orders tables using smart markers | Aspose.Cells smart marker syntax for related tables
// Tags: Aspose.Cells hierarchical smart markers | C# WorkbookDesigner master‑detail report | Excel generation from DataSet Aspose | smart marker template for related tables | master‑detail Excel export Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

namespace MasterDetailReport
{
    // The sample builds an Excel workbook, adds a template with hierarchical smart markers for a Customers master table and its related Orders detail table, creates a DataSet containing both tables and a relation, processes the template with WorkbookDesigner to fill the data, and saves the result as MasterDetailReport.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Report";

            // -------------------------------------------------
            // Build the template with smart markers
            // -------------------------------------------------
            // Header for master (Customer) table
            sheet.Cells["A1"].PutValue("Customer Name");
            sheet.Cells["B1"].PutValue("Customer City");

            // Smart marker row for master records
            // {Customers.Name} and {Customers.City} will be replaced by data
            sheet.Cells["A2"].PutValue("{Customers.Name}");
            sheet.Cells["B2"].PutValue("{Customers.City}");

            // Header for detail (Orders) table - placed below master row
            // We'll leave a blank row for visual separation
            sheet.Cells["A4"].PutValue("Order ID");
            sheet.Cells["B4"].PutValue("Order Date");
            sheet.Cells["C4"].PutValue("Order Amount");

            // Smart marker row for detail records
            // The hierarchical marker {Customers.Orders.OrderID} pulls orders related to each customer
            sheet.Cells["A5"].PutValue("{Customers.Orders.OrderID}");
            sheet.Cells["B5"].PutValue("{Customers.Orders.OrderDate}");
            sheet.Cells["C5"].PutValue("{Customers.Orders.Amount}");

            // -------------------------------------------------
            // Prepare the data source (DataSet with master-detail tables)
            // -------------------------------------------------
            DataSet ds = new DataSet();

            // Master table: Customers
            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("CustomerID", typeof(int));
            customers.Columns.Add("Name", typeof(string));
            customers.Columns.Add("City", typeof(string));
            customers.PrimaryKey = new DataColumn[] { customers.Columns["CustomerID"] };
            ds.Tables.Add(customers);

            // Detail table: Orders
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("CustomerID", typeof(int));
            orders.Columns.Add("OrderDate", typeof(DateTime));
            orders.Columns.Add("Amount", typeof(decimal));
            ds.Tables.Add(orders);

            // Define relation between Customers and Orders
            DataRelation rel = new DataRelation("Customers_Orders",
                customers.Columns["CustomerID"],
                orders.Columns["CustomerID"]);
            ds.Relations.Add(rel);

            // Populate master data
            customers.Rows.Add(1, "Acme Corp", "New York");
            customers.Rows.Add(2, "Beta Ltd", "London");
            customers.Rows.Add(3, "Gamma Inc", "Tokyo");

            // Populate detail data
            orders.Rows.Add(1001, 1, new DateTime(2023, 1, 15), 1250.00m);
            orders.Rows.Add(1002, 1, new DateTime(2023, 2, 20), 980.50m);
            orders.Rows.Add(2001, 2, new DateTime(2023, 3, 5), 430.75m);
            orders.Rows.Add(2002, 2, new DateTime(2023, 4, 12), 2100.00m);
            orders.Rows.Add(3001, 3, new DateTime(2023, 5, 30), 750.00m);

            // -------------------------------------------------
            // Process the smart markers with the data source
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(ds);
            designer.Process(); // Fills the template with master-detail data

            // -------------------------------------------------
            // Save the generated report
            // -------------------------------------------------
            workbook.Save("MasterDetailReport.xlsx", SaveFormat.Xlsx);
        }
    }
}
