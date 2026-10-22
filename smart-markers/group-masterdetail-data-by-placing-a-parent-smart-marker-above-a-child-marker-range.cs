// Title: Create a master‑detail Excel sheet with a parent smart marker above a child range using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that inserts a parent smart marker in cell A1 and a child smart marker block starting at A3:C3, then binds a DataSet with Customers and Orders tables (linked by a DataRelation) to WorkbookDesigner and generates output.xlsx. | Show how to load an existing Excel template or create a new workbook, add hierarchical smart markers for master‑detail data, process them with Aspose.Cells, and save the result. | Demonstrate building a DataSet with related tables, placing smart markers for master and detail rows, and using WorkbookDesigner to produce a populated Excel report.
// Common Searches: aspnet c# how to use Aspose.Cells smart markers for master detail reports | place parent smart marker above child range Aspose.Cells example | generate Excel file with hierarchical data using DataSet relation and smart markers | load template workbook or create new workbook then process smart markers Aspose.Cells | C# Aspose.Cells master‑detail smart markers with DataRelation
// Tags: smart markers for master‑detail hierarchy Aspose.Cells | parent smart marker positioning above child block | WorkbookDesigner hierarchical data processing | C# DataSet relation to Excel via Aspose | template workbook creation Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads or creates a workbook, places a parent smart marker in A1 and a child smart marker range starting at A3:C3, builds a DataSet with Customers and Orders tables linked by a DataRelation, processes the markers with WorkbookDesigner, and saves the populated Excel file as output.xlsx.
class MasterDetailSmartMarkerExample
{
    static void Main()
    {
        try
        {
            // Path to the template workbook
            string templatePath = "template.xlsx";
            Workbook workbook;

            // Load existing template if it exists; otherwise create a new workbook
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                // Ensure there is at least one worksheet
                workbook.Worksheets.Clear();
                workbook.Worksheets.Add("Sheet1");
            }

            // Get the first worksheet where we will place the smart markers
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Insert the parent smart marker (master) above the child range
            // ------------------------------------------------------------
            // Parent marker for the Customer Name in cell A1
            sheet.Cells["A1"].PutValue("&=Customers.Name");

            // ------------------------------------------------------------
            // Insert the child smart marker range (detail) starting at row 2
            // ------------------------------------------------------------
            // Column headers for the child data
            sheet.Cells["A2"].PutValue("Order ID");
            sheet.Cells["B2"].PutValue("Order Date");
            sheet.Cells["C2"].PutValue("Amount");

            // Child smart markers in the row below the headers (row 3)
            sheet.Cells["A3"].PutValue("&=Orders.OrderID");
            sheet.Cells["B3"].PutValue("&=Orders.OrderDate");
            sheet.Cells["C3"].PutValue("&=Orders.Amount");

            // ------------------------------------------------------------
            // Prepare the master‑detail data source (DataSet with relation)
            // ------------------------------------------------------------
            DataSet ds = new DataSet();

            // Master table: Customers
            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("CustomerID", typeof(int));
            customers.Columns.Add("Name", typeof(string));
            customers.Rows.Add(1, "Alice");
            customers.Rows.Add(2, "Bob");
            ds.Tables.Add(customers);

            // Detail table: Orders
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("CustomerID", typeof(int));
            orders.Columns.Add("OrderDate", typeof(DateTime));
            orders.Columns.Add("Amount", typeof(decimal));
            orders.Rows.Add(101, 1, new DateTime(2023, 1, 15), 250.00m);
            orders.Rows.Add(102, 1, new DateTime(2023, 2, 5), 125.50m);
            orders.Rows.Add(201, 2, new DateTime(2023, 3, 12), 300.00m);
            ds.Tables.Add(orders);

            // Define the relation between Customers (parent) and Orders (child)
            DataRelation relation = new DataRelation(
                "CustomerOrders",
                customers.Columns["CustomerID"],
                orders.Columns["CustomerID"]);
            ds.Relations.Add(relation);

            // ------------------------------------------------------------
            // Process the smart markers with the data source
            // ------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(ds);
            designer.Process();

            // ------------------------------------------------------------
            // Save the resulting workbook
            // ------------------------------------------------------------
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
