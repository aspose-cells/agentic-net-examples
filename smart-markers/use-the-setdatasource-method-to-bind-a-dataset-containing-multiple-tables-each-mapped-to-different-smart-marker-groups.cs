using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorkbookDesignerMultipleTablesDemo
    {
        public static void Run()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Prepare a DataSet that contains several DataTables.
                // ------------------------------------------------------------
                DataSet ds = new DataSet();

                // ----- Customers table -------------------------------------------------
                DataTable customers = new DataTable("Customers");
                customers.Columns.Add("CustomerID", typeof(int));
                customers.Columns.Add("Name", typeof(string));
                customers.Columns.Add("City", typeof(string));
                customers.Rows.Add(1, "Alice", "London");
                customers.Rows.Add(2, "Bob", "Paris");
                ds.Tables.Add(customers);

                // ----- Orders table ----------------------------------------------------
                DataTable orders = new DataTable("Orders");
                orders.Columns.Add("OrderID", typeof(int));
                orders.Columns.Add("CustomerID", typeof(int));
                orders.Columns.Add("Product", typeof(string));
                orders.Columns.Add("Quantity", typeof(int));
                orders.Rows.Add(1001, 1, "Laptop", 2);
                orders.Rows.Add(1002, 2, "Smartphone", 5);
                ds.Tables.Add(orders);

                // ------------------------------------------------------------
                // 2. Create a workbook that contains smart markers for each table.
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Smart markers for the Customers table (group name "Customers")
                sheet.Cells["A1"].PutValue("CustomerID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("City");
                sheet.Cells["A2"].PutValue("&=Customers.CustomerID");
                sheet.Cells["B2"].PutValue("&=Customers.Name");
                sheet.Cells["C2"].PutValue("&=Customers.City");

                // Smart markers for the Orders table (group name "Orders")
                // Placed a few rows below the customers block.
                sheet.Cells["A6"].PutValue("OrderID");
                sheet.Cells["B6"].PutValue("CustomerID");
                sheet.Cells["C6"].PutValue("Product");
                sheet.Cells["D6"].PutValue("Quantity");
                sheet.Cells["A7"].PutValue("&=Orders.OrderID");
                sheet.Cells["B7"].PutValue("&=Orders.CustomerID");
                sheet.Cells["C7"].PutValue("&=Orders.Product");
                sheet.Cells["D7"].PutValue("&=Orders.Quantity");

                // ------------------------------------------------------------
                // 3. Bind the DataSet to the designer using SetDataSource(DataSet)
                // ------------------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(ds);

                // ------------------------------------------------------------
                // 4. Process the smart markers and save the result.
                // ------------------------------------------------------------
                designer.Process();

                string outputPath = "MultipleTablesOutput.xlsx";
                // Ensure we can write to the path (overwrite if exists)
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                designer.Workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while generating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }

        // Entry point for the console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}