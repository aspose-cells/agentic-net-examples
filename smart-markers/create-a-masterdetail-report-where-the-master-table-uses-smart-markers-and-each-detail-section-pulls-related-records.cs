// Title: Create a master‑detail Excel report using Aspose.Cells smart markers, DataSet, and DataRelation in C#
// AI Prompts: Design an Excel template with smart markers that auto‑repeat order rows and their corresponding detail rows, then fill it using a C# DataSet and WorkbookDesigner. | Set up Orders and OrderDetails DataTables, link them via a DataRelation on OrderID, bind the DataSet to WorkbookDesigner, process the markers, and export the workbook as MasterDetailReport.xlsx.
// Common Searches: aspocells parent child report smart markers c# example | how to repeat rows for parent and child tables in Excel using Aspose.Cells | populate Excel from DataSet with master‑detail relationship using WorkbookDesigner | specify smart marker range for master‑detail layout in Aspose.Cells | c# generate excel file with orders and order items using smart markers
// Tags: smart markers hierarchical report Aspose.Cells | WorkbookDesigner DataSet binding C# | smart marker range definition Aspose.Cells | repeat rows for related tables Aspose.Cells | DataRelation based Excel report Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

namespace MasterDetailReportExample
{
    // The example builds an Excel template, inserts smart markers to repeat rows for Orders and their related OrderDetails, creates a DataSet with Orders and OrderDetails tables linked by a DataRelation, assigns the DataSet to a WorkbookDesigner, processes all smart markers, and saves the populated workbook as MasterDetailReport.xlsx.
    class Program
    {
        static void Main()
        {
            // ---------- Create template workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Master table header
            cells["A1"].PutValue("Order ID");
            cells["B1"].PutValue("Customer");

            // Start of master table (repeat for each order)
            cells["A2"].PutValue("&=Orders");          // smart marker to repeat rows for Orders table
            cells["A3"].PutValue("&=$OrderID");        // column marker for OrderID
            cells["B3"].PutValue("&=$CustomerName");   // column marker for CustomerName

            // Detail table header (placed after master rows)
            cells["A5"].PutValue("Order Details");
            cells["A6"].PutValue("Product");
            cells["B6"].PutValue("Quantity");

            // Start of detail table (repeat for each related OrderDetails row)
            cells["A7"].PutValue("&=OrderDetails");    // smart marker to repeat rows for OrderDetails table
            cells["A8"].PutValue("&=$Product");        // column marker for Product
            cells["B8"].PutValue("&=$Quantity");       // column marker for Quantity

            // Define the range that contains all smart markers (required when LineByLine = false)
            sheet.Cells.CreateRange("A2:B8").Name = "_CellsSmartMarkers";

            // ---------- Prepare master‑detail data ----------
            DataSet dataSet = new DataSet();

            // Master table: Orders
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("CustomerName", typeof(string));
            orders.Rows.Add(1, "Alice");
            orders.Rows.Add(2, "Bob");

            // Detail table: OrderDetails
            DataTable orderDetails = new DataTable("OrderDetails");
            orderDetails.Columns.Add("OrderID", typeof(int));
            orderDetails.Columns.Add("Product", typeof(string));
            orderDetails.Columns.Add("Quantity", typeof(int));
            orderDetails.Rows.Add(1, "Pen", 10);
            orderDetails.Rows.Add(1, "Notebook", 5);
            orderDetails.Rows.Add(2, "Pencil", 20);
            orderDetails.Rows.Add(2, "Eraser", 15);

            // Add tables to the DataSet
            dataSet.Tables.Add(orders);
            dataSet.Tables.Add(orderDetails);

            // Define relation between master and detail tables
            DataRelation relation = new DataRelation(
                "Orders_OrderDetails",
                orders.Columns["OrderID"],
                orderDetails.Columns["OrderID"]);
            dataSet.Relations.Add(relation);

            // ---------- Process smart markers ----------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            // Use the DataSet as the data source (master‑detail will be resolved automatically)
            designer.SetDataSource(dataSet);

            // Process all smart markers in the defined range
            designer.Process();

            // ---------- Save the result ----------
            designer.Workbook.Save("MasterDetailReport.xlsx");
        }
    }
}
