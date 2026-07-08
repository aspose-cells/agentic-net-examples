using System;
using System.Data;
using Aspose.Cells;

namespace MasterDetailSmartMarkers
{
    class Program
    {
        static void Main()
        {
            // ----- Create a template workbook in memory -----
            Workbook template = new Workbook();
            Worksheet sheet = template.Worksheets[0];
            Cells cells = sheet.Cells;

            // Master table header (smart marker for master records)
            // &="Orders" tells Aspose.Cells to repeat this row for each master row
            cells["A1"].PutValue("&=\"Orders\"");
            cells["A2"].PutValue("Order ID");
            cells["B2"].PutValue("Order Date");
            // Master data row
            cells["A3"].PutValue("&=Orders.OrderID");
            cells["B3"].PutValue("&=Orders.OrderDate");

            // Detail table header (smart marker for child records)
            // &="Orders.Details" repeats this block for each master row's child rows
            cells["A5"].PutValue("&=\"Orders.Details\"");
            cells["A6"].PutValue("Product");
            cells["B6"].PutValue("Quantity");
            // Detail data rows
            cells["A7"].PutValue("&=Orders.Details.ProductName");
            cells["B7"].PutValue("&=Orders.Details.Quantity");

            // ----- Prepare master‑detail data in a DataSet -----
            DataSet ds = new DataSet();

            // Master table
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("OrderDate", typeof(DateTime));
            orders.Rows.Add(1, DateTime.Today.AddDays(-2));
            orders.Rows.Add(2, DateTime.Today.AddDays(-1));
            ds.Tables.Add(orders);

            // Detail table
            DataTable details = new DataTable("Details");
            details.Columns.Add("OrderID", typeof(int)); // foreign key
            details.Columns.Add("ProductName", typeof(string));
            details.Columns.Add("Quantity", typeof(int));
            details.Rows.Add(1, "Apple", 10);
            details.Rows.Add(1, "Banana", 5);
            details.Rows.Add(2, "Orange", 8);
            details.Rows.Add(2, "Grape", 12);
            ds.Tables.Add(details);

            // Define relation between master and detail
            ds.Relations.Add("Orders_Details", orders.Columns["OrderID"], details.Columns["OrderID"]);

            // ----- Process the template with smart markers -----
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = template;
            designer.SetDataSource(ds);
            designer.Process();

            // ----- Save the result -----
            designer.Workbook.Save("MasterDetailReport.xlsx");
        }
    }
}