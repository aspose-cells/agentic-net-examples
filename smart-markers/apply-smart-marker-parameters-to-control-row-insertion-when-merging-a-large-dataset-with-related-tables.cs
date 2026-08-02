// Title: Control Row Insertion in Aspose.Cells Smart Markers (C#) – Using the noadd Parameter with Parent‑Child DataSets
// Description: Demonstrates how to merge a DataSet that contains related tables (Orders and OrderDetails) into an Excel workbook using Aspose.Cells Smart Markers. The example shows the standard "&=Orders" marker for repeating master rows and the "noadd" suffix to suppress automatic row insertion for child rows. It also covers naming the smart‑marker range ("_CellsSmartMarkers") and processing it with WorkbookDesigner.
// Keywords: Aspose.Cells | Smart Markers | noadd parameter | row insertion control | C# | DataSet to Excel | parent‑child tables | WorkbookDesigner | named smart‑marker range | large dataset export
// Common Searches: Aspose.Cells prevent row insertion smart marker | noadd suffix smart markers C# | master detail Excel export Aspose.Cells | process DataSet with related tables using smart markers | named range _CellsSmartMarkers Aspose.Cells
// Developer Intent: The developer needs to decide whether rows should be added for nested smart‑marker tables when a DataSet with parent‑child relations is merged into Excel.
// Use Cases: Create an order report where each order appears once while the detail line stays on a single row to keep the sheet compact. | Export a master‑detail view from a large relational DataSet, disabling row insertion for selected child tables to improve performance. | Generate a printable invoice list where only the header rows repeat and the item rows are displayed without extra row duplication.
// AI Prompts: Show how to apply the noadd suffix in Aspose.Cells smart markers to stop row insertion for a child table. | Provide a C# example that processes a parent‑child DataSet with WorkbookDesigner and a named smart‑marker range. | Explain how to switch between automatic row insertion and noadd behavior for smart markers based on a runtime flag.

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerRowInsertionDemo
{
    // Demonstrates how to merge a DataSet that contains related tables (Orders and OrderDetails) into an Excel workbook using Aspose.Cells Smart Markers. The example shows the standard "&=Orders" marker for repeating master rows and the "noadd" suffix to suppress automatic row insertion for child rows. It also covers naming the smart‑marker range ("_CellsSmartMarkers") and processing it with WorkbookDesigner.
    class Program
    {
        static void Main()
        {
            // ---------- 1. Prepare a DataSet with related tables ----------
            DataSet ds = new DataSet();

            // Orders table
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("Customer", typeof(string));
            orders.Rows.Add(1001, "Alice");
            orders.Rows.Add(1002, "Bob");
            ds.Tables.Add(orders);

            // OrderDetails table (related to Orders)
            DataTable details = new DataTable("OrderDetails");
            details.Columns.Add("OrderID", typeof(int));
            details.Columns.Add("Product", typeof(string));
            details.Columns.Add("Quantity", typeof(int));
            details.Rows.Add(1001, "Laptop", 1);
            details.Rows.Add(1001, "Mouse", 2);
            details.Rows.Add(1002, "Monitor", 1);
            ds.Tables.Add(details);

            // Define relation between Orders and OrderDetails
            ds.Relations.Add("Order_Details",
                orders.Columns["OrderID"],
                details.Columns["OrderID"]);

            // ---------- 2. Create a workbook and place smart markers ----------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Header row for Orders
            cells["A1"].PutValue("Order ID");
            cells["B1"].PutValue("Customer");

            // Smart marker row that will be repeated for each order
            // The marker "&=Orders" tells the designer to start a table for Orders
            cells["A2"].PutValue("&=Orders.OrderID");
            cells["B2"].PutValue("&=Orders.Customer");

            // Header row for OrderDetails (nested table)
            cells["A3"].PutValue("Product");
            cells["B3"].PutValue("Quantity");

            // Smart marker row for details.
            // By default rows are inserted for each detail record.
            // If you want to prevent row insertion, use the "noadd" suffix, e.g. "&=OrderDetails.noadd.Product"
            cells["A4"].PutValue("&=OrderDetails.Product");
            cells["B4"].PutValue("&=OrderDetails.Quantity");

            // Define the range that contains all smart markers.
            // Naming the range "_CellsSmartMarkers" enables range‑based processing.
            ws.Cells.CreateRange("A1:B4").Name = "_CellsSmartMarkers";

            // ---------- 3. Process smart markers with the DataSet ----------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = wb
            };
            designer.SetDataSource(ds);
            designer.Process(); // processes all smart markers in the named range

            // ---------- 4. Save the result ----------
            wb.Save("SmartMarkerRowInsertionResult.xlsx");
        }
    }
}
