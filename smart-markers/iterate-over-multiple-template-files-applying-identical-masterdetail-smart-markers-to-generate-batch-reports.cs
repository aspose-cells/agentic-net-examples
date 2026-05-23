using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace BatchSmartMarkerReport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Folder that contains the template files
                string templateFolder = @"C:\Templates";
                // Folder where the generated reports will be saved
                string outputFolder = @"C:\Reports";

                // Ensure the output folder exists
                Directory.CreateDirectory(outputFolder);

                // List of template file names (add more as needed)
                string[] templateFiles = new string[]
                {
                    Path.Combine(templateFolder, "Template1.xlsx"),
                    Path.Combine(templateFolder, "Template2.xlsx")
                };

                // Prepare a master‑detail DataSet used for every template
                DataSet reportData = CreateMasterDetailDataSet();

                // Iterate over each template, apply smart markers and save the result
                foreach (string templatePath in templateFiles)
                {
                    if (!File.Exists(templatePath))
                    {
                        Console.WriteLine($"Template not found: {templatePath}");
                        continue;
                    }

                    try
                    {
                        // Load the template workbook
                        Workbook workbook = new Workbook(templatePath);

                        // Initialize the designer with the loaded workbook
                        WorkbookDesigner designer = new WorkbookDesigner(workbook);

                        // Assign the same data source to each workbook
                        designer.SetDataSource(reportData);

                        // Process the smart markers (true = preserve unrecognized markers)
                        designer.Process(true);

                        // Build the output file name
                        string outputFileName = Path.GetFileNameWithoutExtension(templatePath) + "_Result.xlsx";
                        string outputPath = Path.Combine(outputFolder, outputFileName);

                        // Save the processed workbook
                        workbook.Save(outputPath);
                        Console.WriteLine($"Report generated: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing template '{templatePath}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Creates a DataSet containing a master table (Orders) and a detail table (OrderDetails)
        // with a relation between them. This DataSet can be reused for all templates.
        private static DataSet CreateMasterDetailDataSet()
        {
            DataSet ds = new DataSet();

            // Master table
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("CustomerName", typeof(string));
            orders.Columns.Add("OrderDate", typeof(DateTime));

            orders.Rows.Add(1, "Alice", new DateTime(2023, 1, 15));
            orders.Rows.Add(2, "Bob", new DateTime(2023, 2, 20));
            orders.Rows.Add(3, "Charlie", new DateTime(2023, 3, 5));

            // Detail table
            DataTable orderDetails = new DataTable("OrderDetails");
            orderDetails.Columns.Add("OrderID", typeof(int));
            orderDetails.Columns.Add("Product", typeof(string));
            orderDetails.Columns.Add("Quantity", typeof(int));
            orderDetails.Columns.Add("UnitPrice", typeof(decimal));

            orderDetails.Rows.Add(1, "Laptop", 1, 1200.00m);
            orderDetails.Rows.Add(1, "Mouse", 2, 25.50m);
            orderDetails.Rows.Add(2, "Desk", 1, 300.00m);
            orderDetails.Rows.Add(3, "Chair", 4, 45.00m);
            orderDetails.Rows.Add(3, "Monitor", 2, 200.00m);

            // Add tables to the DataSet
            ds.Tables.Add(orders);
            ds.Tables.Add(orderDetails);

            // Create a relation between master and detail tables
            ds.Relations.Add("Orders_OrderDetails",
                orders.Columns["OrderID"],
                orderDetails.Columns["OrderID"]);

            return ds;
        }
    }
}