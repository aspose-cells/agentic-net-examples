using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Prepare a DataTable that mimics the result set of a DB query
        // ------------------------------------------------------------
        DataTable dt = new DataTable("Orders");
        dt.Columns.Add("OrderID", typeof(int));
        dt.Columns.Add("CustomerName", typeof(string));
        dt.Columns.Add("OrderDate", typeof(DateTime));
        dt.Columns.Add("Total", typeof(decimal));

        dt.Rows.Add(1001, "Alice",   new DateTime(2023, 1, 15), 250.75m);
        dt.Rows.Add(1002, "Bob",     new DateTime(2023, 2, 20), 120.00m);
        dt.Rows.Add(1003, "Charlie", new DateTime(2023, 3, 5),  560.40m);

        // ------------------------------------------------------------
        // 2. Create a workbook (or load a template that already contains
        //    smart markers). Here we create a simple workbook and add
        //    smart markers manually.
        // ------------------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];

        // Smart markers use the table name ("Orders") followed by column names
        ws.Cells["A1"].PutValue("&=Orders.OrderID");
        ws.Cells["B1"].PutValue("&=Orders.CustomerName");
        ws.Cells["C1"].PutValue("&=Orders.OrderDate");
        ws.Cells["D1"].PutValue("&=Orders.Total");

        // ------------------------------------------------------------
        // 3. Bind the workbook to a WorkbookDesigner
        // ------------------------------------------------------------
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // ------------------------------------------------------------
        // 4. Set the DataTable as the data source.
        //    The DataTable's TableName ("Orders") must match the smart
        //    marker prefix used above.
        // ------------------------------------------------------------
        designer.SetDataSource(dt);

        // ------------------------------------------------------------
        // 5. Process the smart markers – this populates the worksheet
        //    with the data from the DataTable.
        // ------------------------------------------------------------
        designer.Process();

        // ------------------------------------------------------------
        // 6. Save the populated workbook.
        // ------------------------------------------------------------
        designer.Workbook.Save("SmartMarkersFromDataTable.xlsx");
    }
}