using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace BatchSmartMarkerReport
{
    class Program
    {
        static void Main()
        {
            // List of template files to process
            var templates = new List<string>
            {
                "Template1.xlsx",
                "Template2.xlsx",
                "Template3.xlsx"
            };

            // Prepare a DataSet containing master‑detail tables
            DataSet reportData = CreateMasterDetailDataSet();

            // Process each template with the same data source
            foreach (var templatePath in templates)
            {
                // Load the template workbook
                Workbook workbook = new Workbook(templatePath);

                // Initialize the designer with the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Bind the DataSet (master table "Orders", detail table "OrderDetails")
                designer.SetDataSource(reportData);

                // Process all smart markers in the workbook
                designer.Process();

                // Build output file name (e.g., "Template1_Output.xlsx")
                string outputPath = System.IO.Path.GetFileNameWithoutExtension(templatePath) + "_Output.xlsx";

                // Save the processed workbook
                designer.Workbook.Save(outputPath);
                Console.WriteLine($"Processed '{templatePath}' and saved as '{outputPath}'.");
            }
        }

        // Creates a DataSet with two related tables: Orders (master) and OrderDetails (detail)
        private static DataSet CreateMasterDetailDataSet()
        {
            DataSet ds = new DataSet();

            // Master table
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("CustomerName", typeof(string));
            orders.Columns.Add("OrderDate", typeof(DateTime));
            orders.Rows.Add(1001, "Alice", new DateTime(2023, 1, 15));
            orders.Rows.Add(1002, "Bob", new DateTime(2023, 2, 20));
            ds.Tables.Add(orders);

            // Detail table
            DataTable orderDetails = new DataTable("OrderDetails");
            orderDetails.Columns.Add("OrderID", typeof(int));
            orderDetails.Columns.Add("Product", typeof(string));
            orderDetails.Columns.Add("Quantity", typeof(int));
            orderDetails.Columns.Add("UnitPrice", typeof(decimal));
            orderDetails.Rows.Add(1001, "Laptop", 1, 1200.00m);
            orderDetails.Rows.Add(1001, "Mouse", 2, 25.50m);
            orderDetails.Rows.Add(1002, "Desk", 1, 300.00m);
            orderDetails.Rows.Add(1002, "Chair", 4, 45.00m);
            ds.Tables.Add(orderDetails);

            // Define relation for master‑detail processing
            ds.Relations.Add("Order_OrderDetails",
                orders.Columns["OrderID"],
                orderDetails.Columns["OrderID"]);

            return ds;
        }
    }
}