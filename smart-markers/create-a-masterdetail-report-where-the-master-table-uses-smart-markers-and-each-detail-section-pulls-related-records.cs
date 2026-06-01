using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace MasterDetailReportApp
{
    class MasterDetailReport
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a template workbook and place smart markers
                // -------------------------------------------------
                Workbook template = new Workbook();
                Worksheet ws = template.Worksheets[0];

                ws.Cells["A1"].PutValue("Order ID");
                ws.Cells["B1"].PutValue("Order Date");
                ws.Cells["C1"].PutValue("Product");
                ws.Cells["D1"].PutValue("Quantity");

                ws.Cells["A2"].PutValue("&=Orders.OrderID");
                ws.Cells["B2"].PutValue("&=Orders.OrderDate");

                ws.Cells["C3"].PutValue("&=OrderDetails.ProductName");
                ws.Cells["D3"].PutValue("&=OrderDetails.Quantity");

                // Define the range that contains all smart markers.
                // The range must be named "_CellsSmartMarkers" for range smart markers.
                Aspose.Cells.Range smartRange = ws.Cells.CreateRange("A2:D3");
                smartRange.Name = "_CellsSmartMarkers";

                // -------------------------------------------------
                // 2. Prepare master‑detail data in a DataSet
                // -------------------------------------------------
                DataSet ds = new DataSet();

                // Master table: Orders
                DataTable orders = new DataTable("Orders");
                orders.Columns.Add("OrderID", typeof(int));
                orders.Columns.Add("OrderDate", typeof(DateTime));
                orders.Rows.Add(1, DateTime.Today);
                orders.Rows.Add(2, DateTime.Today.AddDays(1));
                ds.Tables.Add(orders);

                // Detail table: OrderDetails
                DataTable orderDetails = new DataTable("OrderDetails");
                orderDetails.Columns.Add("OrderID", typeof(int));          // Foreign key to Orders
                orderDetails.Columns.Add("ProductName", typeof(string));
                orderDetails.Columns.Add("Quantity", typeof(int));
                orderDetails.Rows.Add(1, "Apple", 10);
                orderDetails.Rows.Add(1, "Banana", 5);
                orderDetails.Rows.Add(2, "Orange", 7);
                ds.Tables.Add(orderDetails);

                // Define relation between master and detail tables
                ds.Relations.Add("Order_OrderDetails",
                    orders.Columns["OrderID"]!,
                    orderDetails.Columns["OrderID"]!);

                // -------------------------------------------------
                // 3. Bind the DataSet to the designer and process
                // -------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = template // Assign the template workbook
                };
                designer.SetDataSource(ds); // Set the DataSet as the data source
                designer.Process();         // Populate the smart markers

                // -------------------------------------------------
                // 4. Save the generated report
                // -------------------------------------------------
                string outputPath = "MasterDetailReport.xlsx";
                designer.Workbook.Save(outputPath);
                Console.WriteLine($"Report saved to {Path.GetFullPath(outputPath)}");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"File not found: {ex.FileName}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}