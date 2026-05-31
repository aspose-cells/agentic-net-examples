using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerRowInsertionDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook (template) and add smart markers
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header for parent table "Orders"
            cells["A1"].PutValue("Order ID");
            cells["B1"].PutValue("Customer");

            // Smart markers for parent table rows
            cells["A2"].PutValue("&=$Orders.OrderID");
            cells["B2"].PutValue("&=$Orders.CustomerName");

            // Header for child table "OrderDetails"
            cells["A5"].PutValue("Order Details");
            cells["A6"].PutValue("Product");
            cells["B6"].PutValue("Quantity");

            // Smart markers for child table rows (start below the header)
            cells["A7"].PutValue("&=$OrderDetails.Product");
            cells["B7"].PutValue("&=$OrderDetails.Quantity");

            // -------------------------------------------------
            // 2. Build a DataSet with related tables
            // -------------------------------------------------
            DataSet ds = new DataSet();

            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("CustomerName", typeof(string));
            orders.Rows.Add(1001, "Alice");
            orders.Rows.Add(1002, "Bob");
            ds.Tables.Add(orders);

            DataTable details = new DataTable("OrderDetails");
            details.Columns.Add("OrderID", typeof(int));
            details.Columns.Add("Product", typeof(string));
            details.Columns.Add("Quantity", typeof(int));
            details.Rows.Add(1001, "Laptop", 1);
            details.Rows.Add(1001, "Mouse", 2);
            details.Rows.Add(1002, "Monitor", 1);
            ds.Tables.Add(details);

            // -------------------------------------------------
            // 3. Bind the DataSet to WorkbookDesigner and process smart markers
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                LineByLine = false // Use range‑based processing (optional)
            };
            designer.SetDataSource(ds);
            designer.Process();

            // -------------------------------------------------
            // 4. Control row insertion for the child table using ImportTableOptions
            // -------------------------------------------------
            ImportTableOptions importOptions = new ImportTableOptions
            {
                InsertRows = true,          // Insert new rows instead of overwriting existing ones
                IsFieldNameShown = false,   // Field names are already present in the template
                ShiftFirstRowDown = true    // Keep the header row (A6:B6) intact
            };

            // Starting cell of the child smart‑marker block (zero‑based indices)
            int startRow = 6;    // Row 7 in Excel (A7)
            int startColumn = 0; // Column A

            // Re‑import the OrderDetails table with the above options
            cells.ImportData(details, startRow, startColumn, importOptions);

            // -------------------------------------------------
            // 5. Save the populated workbook
            // -------------------------------------------------
            workbook.Save("SmartMarkerRowInsertionResult.xlsx");
        }
    }
}