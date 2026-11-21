// Title: Create separate worksheets for each master‑detail group using Aspose.Cells smart markers with sheet‑break syntax in C#
// Common Searches: C# Aspose.Cells smart marker sheet break to generate a worksheet for each order | How to split master‑detail data into multiple Excel sheets using Aspose.Cells WorkbookDesigner | Using sheet break marker in an Excel template to create separate sheets per master record with Aspose.Cells | Aspose.Cells example for grouping orders and order details into individual worksheets | Process DataSet with relations in Aspose.Cells smart markers to produce per‑group worksheets
// Tags: WorkbookDesigner dataset integration | Aspose.Cells worksheet break syntax | master‑detail worksheet generation C# | Excel smart marker group output | per‑order worksheet creation Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

// The example loads a template workbook containing a sheet‑break smart marker, builds master (Orders) and detail (OrderDetails) DataTables, defines a relation on OrderID, sets the combined DataSet as the data source for WorkbookDesigner, processes the smart markers to create an individual worksheet for each OrderID, and saves the result to Result.xlsx.
class Program
{
    static void Main()
    {
        // Load the template workbook that contains smart markers.
        // The template should have a sheet break marker like &="SheetBreak"
        // placed before the master‑detail group to create a new worksheet per group.
        Workbook workbook = new Workbook("Template.xlsx");

        // ---------- Prepare master‑detail data ----------
        // Master table (e.g., Orders)
        DataTable master = new DataTable("Orders");
        master.Columns.Add("OrderID", typeof(int));
        master.Columns.Add("CustomerName", typeof(string));
        master.Rows.Add(1, "Alice");
        master.Rows.Add(2, "Bob");
        master.Rows.Add(3, "Charlie");

        // Detail table (e.g., OrderDetails)
        DataTable detail = new DataTable("OrderDetails");
        detail.Columns.Add("OrderID", typeof(int));
        detail.Columns.Add("Product", typeof(string));
        detail.Columns.Add("Quantity", typeof(int));
        detail.Rows.Add(1, "Apple", 10);
        detail.Rows.Add(1, "Banana", 5);
        detail.Rows.Add(2, "Orange", 7);
        detail.Rows.Add(2, "Grape", 3);
        detail.Rows.Add(3, "Mango", 12);
        detail.Rows.Add(3, "Peach", 8);

        // Combine tables into a DataSet and define the relation.
        DataSet ds = new DataSet();
        ds.Tables.Add(master);
        ds.Tables.Add(detail);
        ds.Relations.Add("Orders_OrderDetails",
                         master.Columns["OrderID"],
                         detail.Columns["OrderID"]);

        // ---------- Process smart markers ----------
        // WorkbookDesigner handles smart marker processing.
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(ds);
        designer.Process(); // This will create a separate worksheet for each OrderID.

        // ---------- Save the output ----------
        workbook.Save("Result.xlsx");
    }
}
