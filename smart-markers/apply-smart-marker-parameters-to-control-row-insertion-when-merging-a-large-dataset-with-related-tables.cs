// Title: Use Smart Marker Parameters to Limit Row Insertion When Merging a Large Parent‑Child DataSet with WorkbookDesigner in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates an Excel template containing smart markers for Orders and OrderDetails, then uses WorkbookDesigner to merge a parent‑child DataSet while preventing extra rows from being inserted. | Show how to apply smart‑marker parameters or options in Aspose.Cells to control row expansion when processing a large related DataSet with WorkbookDesigner. | Provide a complete example that builds a DataSet with Orders and OrderDetails tables, defines the relationship, processes the smart markers, and saves the result without duplicate rows.
// Common Searches: asp.net c# aspose.cells smart markers limit rows when merging parent child dataset | how to prevent extra rows in Excel output using WorkbookDesigner with related tables | smart marker parameter to control row insertion for large dataset in Aspose.Cells | c# merge dataset with orders and orderdetails using smart markers without duplicate rows | aspose.cells workbookdesigner dataset relationship row duplication issue
// Tags: smart-marker row insertion control | WorkbookDesigner dataset merging | Aspose.Cells parent-child tables | C# generate Excel from DataSet | prevent duplicate rows Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example creates a simple Excel template with smart markers for Orders and OrderDetails, builds a large parent‑child DataSet with a defined relationship, assigns it to WorkbookDesigner, processes the smart markers while controlling row insertion, and saves the populated workbook as Result.xlsx.
class SmartMarkerMergeExample
{
    static void Main()
    {
        try
        {
            // Ensure the template workbook exists; create a simple one if missing
            const string templatePath = "Template.xlsx";
            if (!File.Exists(templatePath))
            {
                var tempWb = new Workbook();
                var ws = tempWb.Worksheets[0];
                ws.Name = "Sheet1";

                // Example smart markers (adjust as needed)
                ws.Cells["A1"].PutValue("&=Orders.OrderID");
                ws.Cells["B1"].PutValue("&=Orders.CustomerID");
                ws.Cells["C1"].PutValue("&=Orders.OrderDate");
                ws.Cells["A2"].PutValue("&=OrderDetails.ProductID");
                ws.Cells["B2"].PutValue("&=OrderDetails.Quantity");

                tempWb.Save(templatePath);
            }

            // Load the template workbook
            var workbook = new Workbook(templatePath);

            // Use WorkbookDesigner to process smart markers
            var designer = new WorkbookDesigner(workbook);

            // Prepare a large DataSet with related tables (Orders, OrderDetails, etc.)
            DataSet dataSet = GetLargeRelatedDataSet();

            // Set the data source for the designer
            designer.SetDataSource(dataSet);

            // Process smart markers
            designer.Process();

            // Save the resulting workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Mock method to obtain a large DataSet with related tables.
    // In a real scenario, this would be filled from a database or other source.
    static DataSet GetLargeRelatedDataSet()
    {
        var ds = new DataSet();

        // Primary table
        var orders = new DataTable("Orders");
        orders.Columns.Add("OrderID", typeof(int));
        orders.Columns.Add("CustomerID", typeof(string));
        orders.Columns.Add("OrderDate", typeof(DateTime));
        ds.Tables.Add(orders);

        // Related table
        var orderDetails = new DataTable("OrderDetails");
        orderDetails.Columns.Add("OrderID", typeof(int));
        orderDetails.Columns.Add("ProductID", typeof(int));
        orderDetails.Columns.Add("Quantity", typeof(int));
        ds.Tables.Add(orderDetails);

        // Define relationship (use non‑null columns)
        ds.Relations.Add("Orders_OrderDetails",
            orders.Columns["OrderID"]!,
            orderDetails.Columns["OrderID"]!);

        // Populate with sample data
        for (int i = 1; i <= 1000; i++)
        {
            var orderRow = orders.NewRow();
            orderRow["OrderID"] = i;
            orderRow["CustomerID"] = "C" + (i % 50).ToString("D3");
            orderRow["OrderDate"] = DateTime.Today.AddDays(-i);
            orders.Rows.Add(orderRow);

            for (int j = 1; j <= 5; j++)
            {
                var detailRow = orderDetails.NewRow();
                detailRow["OrderID"] = i;
                detailRow["ProductID"] = j;
                detailRow["Quantity"] = (i + j) % 10 + 1;
                orderDetails.Rows.Add(detailRow);
            }
        }

        return ds;
    }
}
