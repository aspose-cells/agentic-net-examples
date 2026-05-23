using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDetailExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- Create a new workbook --------------------
                Workbook workbook = new Workbook();

                // -------------------- Create Master sheet with smart markers --------------------
                Worksheet masterSheet = workbook.Worksheets[0];
                masterSheet.Name = "Master";

                // Header row
                masterSheet.Cells["A1"].PutValue("OrderID");
                masterSheet.Cells["B1"].PutValue("OrderDate");

                // Master data smart markers
                masterSheet.Cells["A2"].PutValue("&=Orders.OrderID");
                masterSheet.Cells["B2"].PutValue("&=Orders.OrderDate");

                // -------------------- Create Detail sheet with its own smart markers --------------------
                Worksheet detailSheet = workbook.Worksheets.Add("Detail");

                // Header row for detail data
                detailSheet.Cells["A1"].PutValue("Product");
                detailSheet.Cells["B1"].PutValue("Quantity");

                // Detail rows smart markers (will be populated from the detail data source)
                detailSheet.Cells["A2"].PutValue("&=OrdersDetails.Product");
                detailSheet.Cells["B2"].PutValue("&=OrdersDetails.Quantity");

                // -------------------- Prepare data sources --------------------
                // Master data table
                DataTable ordersTable = new DataTable("Orders");
                ordersTable.Columns.Add("OrderID", typeof(int));
                ordersTable.Columns.Add("OrderDate", typeof(DateTime));
                ordersTable.Rows.Add(1001, DateTime.Now);
                ordersTable.Rows.Add(1002, DateTime.Now.AddDays(1));

                // Detail data table (multiple rows per master record)
                DataTable detailsTable = new DataTable("OrdersDetails");
                detailsTable.Columns.Add("Product", typeof(string));
                detailsTable.Columns.Add("Quantity", typeof(int));
                detailsTable.Rows.Add("Apple", 10);
                detailsTable.Rows.Add("Banana", 20);
                detailsTable.Rows.Add("Cherry", 15);

                // -------------------- Set up WorkbookDesigner --------------------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Bind data sources
                designer.SetDataSource("Orders", ordersTable);
                designer.SetDataSource("OrdersDetails", detailsTable);

                // Process smart markers (no need for DetailSheet property in recent API versions)
                designer.Process();

                // -------------------- Save the result --------------------
                string outputPath = "SmartMarkerDetailOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}