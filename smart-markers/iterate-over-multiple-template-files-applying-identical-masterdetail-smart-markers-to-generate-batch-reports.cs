// Title: Batch generate Excel reports from multiple .xlsx templates using master‑detail smart markers with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells process all .xlsx templates in a folder with master‑detail smart markers | How to use WorkbookDesigner to apply the same smart markers to multiple Excel files in .NET | Generate batch reports from several Excel templates using a DataSet and smart markers in Aspose.Cells | Iterate over Excel template files and create reports with master‑detail smart markers in C# | Aspose.Cells batch report generation from multiple templates using WorkbookDesigner and DataSet
// Tags: batch report generation WorkbookDesigner | process multiple Excel templates Aspose.Cells | C# load and save .xlsx files Aspose.Cells | iterate template folder Aspose.Cells | apply smart markers to Excel workbooks C#

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace BatchReportGenerator
{
    // The example scans a directory for all *.xlsx template files, loads each workbook, assigns a sample DataSet containing Orders and OrderDetails tables as the data source for master‑detail smart markers via WorkbookDesigner, processes the markers, and saves each generated report with a '_Report.xlsx' suffix to an output folder, handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Folder containing the template files (e.g., *.xlsx)
                string templateFolder = @"C:\Templates";
                // Folder where the generated reports will be saved
                string outputFolder = @"C:\Reports";

                // Verify template folder exists
                if (!Directory.Exists(templateFolder))
                {
                    Console.WriteLine($"Template folder does not exist: {templateFolder}");
                    return;
                }

                // Ensure output folder exists
                Directory.CreateDirectory(outputFolder);

                // Get all Excel template files in the folder
                string[] templateFiles = Directory.GetFiles(templateFolder, "*.xlsx");

                if (templateFiles.Length == 0)
                {
                    Console.WriteLine("No template files found.");
                    return;
                }

                // Iterate over each template file
                foreach (string templatePath in templateFiles)
                {
                    try
                    {
                        // Verify the template file exists before loading
                        if (!File.Exists(templatePath))
                        {
                            Console.WriteLine($"Template file not found: {templatePath}");
                            continue;
                        }

                        // Load the template workbook
                        Workbook workbook = new Workbook(templatePath);

                        // Prepare the data source for master‑detail smart markers
                        DataSet dataSet = GetSampleDataSet();

                        // Use WorkbookDesigner (compatible with all Aspose.Cells versions) to process smart markers
                        WorkbookDesigner designer = new WorkbookDesigner(workbook);
                        designer.SetDataSource(dataSet);
                        designer.Process();

                        // Build the output file name based on the template name
                        string templateFileName = Path.GetFileNameWithoutExtension(templatePath);
                        string outputPath = Path.Combine(outputFolder, $"{templateFileName}_Report.xlsx");

                        // Save the generated report
                        workbook.Save(outputPath, SaveFormat.Xlsx);

                        Console.WriteLine($"Generated report: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing template '{templatePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch report generation completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        /// <returns>DataSet with "Orders" and "OrderDetails" tables.</returns>
        private static DataSet GetSampleDataSet()
        {
            DataSet ds = new DataSet();

            // Master table: Orders
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("CustomerName", typeof(string));
            orders.Columns.Add("OrderDate", typeof(DateTime));

            // Add sample master rows
            orders.Rows.Add(1, "Alice", new DateTime(2023, 5, 1));
            orders.Rows.Add(2, "Bob", new DateTime(2023, 5, 3));
            orders.Rows.Add(3, "Charlie", new DateTime(2023, 5, 5));

            // Detail table: OrderDetails
            DataTable orderDetails = new DataTable("OrderDetails");
            orderDetails.Columns.Add("OrderID", typeof(int));
            orderDetails.Columns.Add("Product", typeof(string));
            orderDetails.Columns.Add("Quantity", typeof(int));
            orderDetails.Columns.Add("UnitPrice", typeof(decimal));

            // Add sample detail rows
            orderDetails.Rows.Add(1, "Laptop", 1, 1200.00m);
            orderDetails.Rows.Add(1, "Mouse", 2, 25.00m);
            orderDetails.Rows.Add(2, "Desk", 1, 300.00m);
            orderDetails.Rows.Add(3, "Chair", 4, 45.00m);
            orderDetails.Rows.Add(3, "Monitor", 2, 200.00m);

            // Add tables to the DataSet
            ds.Tables.Add(orders);
            ds.Tables.Add(orderDetails);

            // Define the relation between master and detail tables
            ds.Relations.Add(
                "Order_OrderDetails",
                orders.Columns["OrderID"]!,
                orderDetails.Columns["OrderID"]!);

            return ds;
        }
    }
}
