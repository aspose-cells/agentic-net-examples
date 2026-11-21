// Title: Using the InsertRows smart marker parameter in Aspose.Cells C# to merge a master‑detail DataSet into an Excel report
// Common Searches: C# example for inserting rows with smart markers when exporting a DataSet to Excel | Aspose.Cells master‑detail report using WorkbookDesigner and InsertRows | How to repeat Excel rows for each related record using smart markers in C# | Generate Excel file from related DataTables with row duplication via smart markers | Using Aspose.Cells smart marker parameters to control row insertion in a template
// Tags: InsertRows smart marker Aspose.Cells C# | master‑detail DataSet to Excel WorkbookDesigner | smart marker row duplication Excel template | Aspose.Cells data source relation handling | C# generate Excel report from related tables

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerExample
{
    // The sample loads or creates an Excel template, builds a master‑detail DataSet (Customers and Orders), defines a relation, places a smart marker with the InsertRows parameter to repeat rows for each order, processes the markers via WorkbookDesigner, and saves the populated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string resultPath = "Result.xlsx";

                // Load the template workbook if it exists; otherwise create a new workbook.
                Workbook workbook;
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    // Create a simple workbook with a placeholder smart marker row.
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "Report";

                    // Example smart marker row: {Customers.CustomerName} {Orders.OrderID}
                    // InsertRows parameter will repeat this row for each order.
                    sheet.Cells["A1"].PutValue("{Customers.CustomerName}");
                    sheet.Cells["B1"].PutValue("{Orders.OrderID, InsertRows}");
                }

                // Build the DataSet with master‑detail tables.
                DataSet dataSet = new DataSet();

                // Master table: Customers
                DataTable customers = new DataTable("Customers");
                customers.Columns.Add("CustomerID", typeof(int));
                customers.Columns.Add("CustomerName", typeof(string));
                customers.Rows.Add(1, "Acme Corp");
                customers.Rows.Add(2, "Globex Inc");
                dataSet.Tables.Add(customers);

                // Detail table: Orders (related to Customers by CustomerID)
                DataTable orders = new DataTable("Orders");
                orders.Columns.Add("OrderID", typeof(int));
                orders.Columns.Add("CustomerID", typeof(int));
                orders.Columns.Add("OrderDate", typeof(DateTime));
                orders.Columns.Add("Amount", typeof(decimal));
                orders.Rows.Add(101, 1, new DateTime(2023, 1, 15), 1250.00m);
                orders.Rows.Add(102, 1, new DateTime(2023, 2, 3), 980.50m);
                orders.Rows.Add(201, 2, new DateTime(2023, 1, 22), 760.00m);
                dataSet.Tables.Add(orders);

                // Define relation for master‑detail navigation.
                dataSet.Relations.Add(
                    "Customer_Orders",
                    customers.Columns["CustomerID"]!,
                    orders.Columns["CustomerID"]!);

                // Process smart markers using WorkbookDesigner.
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSet);
                designer.Process();

                // Save the result.
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
