// Title: How to control row insertion with the 'noadd' smart marker parameter when merging parent‑child DataSet tables using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that applies the 'noadd' smart marker parameter to child table fields to stop automatic row insertion while processing a DataSet with Aspose.Cells. | Show how to manually import rows of a child DataTable into an Excel worksheet after smart marker processing by using ImportData with InsertRows enabled. | Demonstrate defining a named smart‑marker range (e.g., _CellsSmartMarkers) so that WorkbookDesigner processes only the specified cells. | Explain how to bind a parent‑child DataSet to WorkbookDesigner and generate the final workbook with mixed automatic and manual row insertion.
// Common Searches: Aspose.Cells C# smart marker noadd parameter for child table rows | prevent automatic row insertion when using smart markers with related tables | import child DataTable rows after smart marker processing Aspose.Cells | use named range _CellsSmartMarkers with WorkbookDesigner | merge parent child DataSet into Excel using smart markers and manual row insertion
// Tags: Aspose.Cells smart marker noadd control | C# parent child DataSet Excel export | WorkbookDesigner named smart marker range | ImportData InsertRows option example | manual row insertion after smart marker processing

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerRowInsertionDemo
{
    // The example creates a DataSet with Customers and Orders tables linked by a relation, builds an Excel template with smart markers where the CustomerName uses default row insertion and the Orders fields use the 'noadd' parameter to suppress automatic rows, processes the markers via WorkbookDesigner, then manually inserts the Orders rows using ImportData with InsertRows enabled, and finally saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- 1. Prepare a DataSet with related tables ----------
                DataSet ds = new DataSet();

                // Parent table: Customers
                DataTable customers = new DataTable("Customers");
                customers.Columns.Add("CustomerID", typeof(int));
                customers.Columns.Add("CustomerName", typeof(string));
                customers.Rows.Add(1, "Alpha Corp");
                customers.Rows.Add(2, "Beta Ltd");
                ds.Tables.Add(customers);

                // Child table: Orders (related to Customers)
                DataTable orders = new DataTable("Orders");
                orders.Columns.Add("OrderID", typeof(int));
                orders.Columns.Add("CustomerID", typeof(int));
                orders.Columns.Add("Product", typeof(string));
                orders.Columns.Add("Quantity", typeof(int));
                orders.Rows.Add(1001, 1, "Laptop", 5);
                orders.Rows.Add(1002, 1, "Mouse", 10);
                orders.Rows.Add(2001, 2, "Monitor", 3);
                ds.Tables.Add(orders);

                // Define relation between Customers and Orders
                ds.Relations.Add(
                    "CustOrders",
                    customers.Columns["CustomerID"]!,
                    orders.Columns["CustomerID"]!
                );

                // ---------- 2. Create a workbook template with smart markers ----------
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                Cells cells = ws.Cells;

                // Header row
                cells["A1"].PutValue("Customer");
                cells["B1"].PutValue("Order ID");
                cells["C1"].PutValue("Product");
                cells["D1"].PutValue("Quantity");

                // Smart marker rows (starting at row 2)
                // Parent table: let Aspose.Cells insert rows automatically.
                cells["A2"].PutValue("&=Customers.CustomerName");
                // Child table: use 'noadd' to prevent automatic row insertion.
                cells["B2"].PutValue("&=Orders.OrderID,noadd");
                cells["C2"].PutValue("&=Orders.Product,noadd");
                cells["D2"].PutValue("&=Orders.Quantity,noadd");

                // Define the range that contains smart markers.
                // Naming the range "_CellsSmartMarkers" tells the designer to process only this range.
                AsposeRange smRange = cells.CreateRange("A2:D2");
                smRange.Name = "_CellsSmartMarkers";

                // ---------- 3. Bind the DataSet to the designer ----------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = wb
                };
                designer.SetDataSource(ds);

                // ---------- 4. Process smart markers ----------
                // Only the customer names are populated; order fields stay empty because of 'noadd'.
                designer.Process();

                // ---------- 5. Manually insert rows for the child table using ImportData ----------
                ImportTableOptions importOpts = new ImportTableOptions
                {
                    InsertRows = true,          // Insert new rows instead of overwriting.
                    IsFieldNameShown = false,   // Do not import column names.
                    ShiftFirstRowDown = true    // Keep the header row intact.
                };

                // Import the Orders table starting just below the first data row (row index 1, column index 1).
                cells.ImportData(orders, 1, 1, importOpts);

                // ---------- 6. Save the result ----------
                string outputPath = "SmartMarkerRowInsertionResult.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
