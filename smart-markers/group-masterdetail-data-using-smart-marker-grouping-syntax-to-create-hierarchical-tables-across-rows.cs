using System;
using System.Data;
using System.IO;
using Aspose.Cells;

class MasterDetailSmartMarkerDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ----- Define the template with smart markers -----
            // Master (order) header
            cells["A1"].PutValue("Order ID");
            cells["B1"].PutValue("Order Date");

            // Master data markers
            cells["A2"].PutValue("&=Orders.OrderID");
            cells["B2"].PutValue("&=Orders.OrderDate");

            // Detail (order details) header
            cells["A3"].PutValue("Product");
            cells["B3"].PutValue("Quantity");

            // Detail data markers using grouping syntax
            cells["A4"].PutValue("&=Orders.OrderDetails.Product");
            cells["B4"].PutValue("&=Orders.OrderDetails.Quantity");

            // Mark the range that contains the smart markers
            worksheet.Cells.CreateRange("A1:B4").Name = "_CellsSmartMarkers";

            // ----- Prepare master‑detail data in a DataSet -----
            DataSet dataSet = new DataSet();

            // Master table: Orders
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("OrderDate", typeof(DateTime));
            orders.Rows.Add(1, new DateTime(2023, 1, 10));
            orders.Rows.Add(2, new DateTime(2023, 2, 15));
            dataSet.Tables.Add(orders);

            // Detail table: OrderDetails
            DataTable orderDetails = new DataTable("OrderDetails");
            orderDetails.Columns.Add("OrderID", typeof(int));
            orderDetails.Columns.Add("Product", typeof(string));
            orderDetails.Columns.Add("Quantity", typeof(int));
            orderDetails.Rows.Add(1, "Apple", 10);
            orderDetails.Rows.Add(1, "Banana", 5);
            orderDetails.Rows.Add(2, "Orange", 8);
            orderDetails.Rows.Add(2, "Grape", 12);
            dataSet.Tables.Add(orderDetails);

            // Define the relation between master and detail
            dataSet.Relations.Add("Order_OrderDetails",
                orders.Columns["OrderID"],
                orderDetails.Columns["OrderID"]);

            // ----- Process the smart markers -----
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // ----- Optional: group each master row with its detail rows -----
            // After processing, each master row is followed by its detail rows.
            int currentRow = 1; // zero‑based index of the first master row (A2)
            foreach (DataRow order in orders.Rows)
            {
                // Count how many detail rows belong to this order
                int detailCount = orderDetails.Select($"OrderID = {order["OrderID"]}").Length;

                if (detailCount > 0)
                {
                    // Group the master row and its detail rows (inclusive)
                    int firstRow = currentRow;
                    int lastRow = currentRow + detailCount;
                    cells.GroupRows(firstRow, lastRow, false);
                }

                // Move to the next master row position
                currentRow += detailCount + 1;
            }

            // Save the resulting workbook (ensure the directory exists)
            string outputPath = "MasterDetailSmartMarkers.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}