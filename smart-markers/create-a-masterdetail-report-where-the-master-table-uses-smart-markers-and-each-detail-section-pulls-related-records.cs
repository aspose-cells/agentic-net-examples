// Title: C# Example: Master‑Detail Excel Report Using Aspose.Cells Smart Markers
// Description: This self‑contained C# sample creates a DataSet with Orders and OrderDetails tables, defines a relation on OrderID, builds a workbook template with smart markers, marks the range as "_CellsSmartMarkers", binds the DataSet to a WorkbookDesigner, processes the markers to repeat detail rows for each order, and saves the result as MasterDetailReport.xlsx.
// Keywords: Aspose.Cells | C# | smart markers | master detail report | Excel generation | WorkbookDesigner | DataSet relation | _CellsSmartMarkers | GitHub example | code sample
// Common Searches: Aspose.Cells master detail smart markers C# | How to use _CellsSmartMarkers range | Create Excel report from DataSet with relations | C# Aspose.Cells example for master‑detail | Generate Excel with repeating rows using smart markers
// Developer Intent: Generate an Excel file that lists each order followed by its related line items using smart markers.
// Use Cases: Automated invoice creation where each invoice header is followed by its item rows. | Sales dashboards that group product quantities under each customer order without manual loops. | Exporting relational database query results to a formatted Excel workbook with hierarchical layout.
// AI Prompts: Add a subtotal row for each order in the smart‑marker template. | Include a second detail table (e.g., shipments) in the same master‑detail report. | Insert a page break after every master record when generating the Excel file.

using System;
using System.Data;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace MasterDetailSmartMarkersDemo
{
    // This self‑contained C# sample creates a DataSet with Orders and OrderDetails tables, defines a relation on OrderID, builds a workbook template with smart markers, marks the range as "_CellsSmartMarkers", binds the DataSet to a WorkbookDesigner, processes the markers to repeat detail rows for each order, and saves the result as MasterDetailReport.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- 1. Prepare master‑detail data ----------
                DataSet ds = new DataSet();

                // Master table: Orders
                DataTable orders = new DataTable("Orders");
                orders.Columns.Add("OrderID", typeof(int));
                orders.Columns.Add("CustomerName", typeof(string));
                orders.Rows.Add(1, "Alice");
                orders.Rows.Add(2, "Bob");
                ds.Tables.Add(orders);

                // Detail table: OrderDetails
                DataTable orderDetails = new DataTable("OrderDetails");
                orderDetails.Columns.Add("OrderID", typeof(int));
                orderDetails.Columns.Add("Product", typeof(string));
                orderDetails.Columns.Add("Quantity", typeof(int));
                orderDetails.Rows.Add(1, "Apple", 10);
                orderDetails.Rows.Add(1, "Banana", 5);
                orderDetails.Rows.Add(2, "Orange", 7);
                orderDetails.Rows.Add(2, "Grape", 3);
                ds.Tables.Add(orderDetails);

                // Define relation between master and detail
                ds.Relations.Add(
                    "Orders_Details",
                    orders.Columns["OrderID"]!,
                    orderDetails.Columns["OrderID"]!);

                // ---------- 2. Build a template workbook with smart markers ----------
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                Cells cells = ws.Cells;

                // Header row
                cells["A1"].PutValue("Order ID");
                cells["B1"].PutValue("Customer");
                cells["C1"].PutValue("Product");
                cells["D1"].PutValue("Quantity");

                // Master smart markers (first two columns)
                cells["A2"].PutValue("&=Orders.OrderID");
                cells["B2"].PutValue("&=Orders.CustomerName");

                // Detail smart markers (next two columns)
                // The detail markers are placed in the same row; Aspose will repeat the row for each related detail record.
                cells["C2"].PutValue("&=OrderDetails.Product");
                cells["D2"].PutValue("&=OrderDetails.Quantity");

                // Define the range that contains smart markers.
                // Naming the range as "_CellsSmartMarkers" tells the designer to process it as a block.
                AsposeRange smartRange = cells.CreateRange("A2:D2");
                smartRange.Name = "_CellsSmartMarkers";

                // ---------- 3. Process the template with the data source ----------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = wb // assign the template workbook
                };
                designer.SetDataSource(ds); // bind the DataSet (master + detail)
                designer.Process();         // populate smart markers

                // ---------- 4. Save the result ----------
                wb.Save("MasterDetailReport.xlsx");
                Console.WriteLine("Report generated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
