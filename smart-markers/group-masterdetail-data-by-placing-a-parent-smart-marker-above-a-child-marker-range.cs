// Title: Generate a master‑detail Excel file with parent and child smart markers using Aspose.Cells for .NET (C#)
// Common Searches: how to create a master detail report with smart markers in Aspose.Cells C# | Aspose.Cells example of placing a parent smart marker above child markers | C# generate hierarchical Excel using WorkbookDesigner and a DataSet | group orders by customer using smart markers in Aspose.Cells .NET | using parent‑child smart markers to produce Excel output with Aspose.Cells
// Tags: smart markers hierarchical data binding | Aspose.Cells parent child marker layout | C# dataset binding with WorkbookDesigner | Excel generation using smart markers | group related tables in Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

// The sample creates a workbook, adds a parent smart marker (&=Customers.Name) in cell A1 and child markers (&=Orders.OrderID, &=Orders.OrderDate) in row 2, builds a DataSet with Customers and Orders tables linked by CustomerName, assigns the DataSet to a WorkbookDesigner, processes the smart markers to produce a master‑detail layout, and saves the result as MasterDetailSmartMarkers.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Smart marker layout -----
            // Parent marker: will repeat for each customer
            sheet.Cells["A1"].PutValue("&=Customers.Name");

            // Child markers: will repeat for each order belonging to the current customer
            sheet.Cells["A2"].PutValue("&=Orders.OrderID");
            sheet.Cells["B2"].PutValue("&=Orders.OrderDate");

            // ----- Sample data -----
            // Master table: Customers
            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("Name", typeof(string));
            customers.Rows.Add("Alice");
            customers.Rows.Add("Bob");

            // Detail table: Orders (includes a foreign key to link to Customers)
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("OrderDate", typeof(string));
            orders.Columns.Add("CustomerName", typeof(string)); // link column
            orders.Rows.Add(1, "2023-01-01", "Alice");
            orders.Rows.Add(2, "2023-01-05", "Alice");
            orders.Rows.Add(3, "2023-02-10", "Bob");

            // Combine tables into a DataSet
            DataSet ds = new DataSet();
            ds.Tables.Add(customers);
            ds.Tables.Add(orders);

            // Process the smart markers using WorkbookDesigner (handles master‑detail automatically)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(ds);   // set the data source
            designer.Process();           // generate the output

            // Define output file name
            string outputPath = "MasterDetailSmartMarkers.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
