// Title: Batch generate Excel reports from multiple templates using master‑detail smart markers with Aspose.Cells for .NET (C#)
// Description: Loads a collection of Excel templates that share identical master‑detail smart markers, creates a DataSet with Orders and OrderDetails tables, binds them to a WorkbookDesigner, processes the markers while preserving unknown tags, and saves each populated workbook to a designated output folder.
// Keywords: Aspose.Cells | C# | .NET | smart markers | master detail | batch report generation | Excel template processing | WorkbookDesigner | DataSet binding | multiple workbooks | automated Excel reports
// Common Searches: Aspose.Cells process multiple Excel templates C# | master‑detail smart markers batch generation | how to use WorkbookDesigner with several workbooks | generate Excel reports from a list of templates .NET | apply same smart markers to many Excel files | C# batch Excel report automation Aspose
// Developer Intent: Create a set of Excel reports by applying identical master‑detail smart markers to many template files in a single run.
// Use Cases: Iterate over a list of template paths, load each workbook, bind Orders and OrderDetails tables, process smart markers, and save the result with a unique filename. | Validate template existence before processing to prevent runtime errors and log missing files. | Configure WorkbookDesigner.LineByLine = false when markers reside in named ranges to keep the original layout. | Extend the DataSet with additional detail tables (e.g., ShippingInfo) and bind them to the designer for richer reports. | Integrate logging or telemetry to monitor batch processing performance across large numbers of templates.
// AI Prompts: Generate C# code that reads an array of Excel template files and uses Aspose.Cells WorkbookDesigner to apply master‑detail smart markers from a DataSet, saving each output with a custom naming pattern. | Explain how to add error‑handling, logging, and retry logic for batch processing of Excel templates with Aspose.Cells smart markers. | Suggest ways to modify the example to support dynamic template discovery from a directory and parallel processing for faster report generation.

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace BatchSmartMarkerReport
{
    // Loads a collection of Excel templates that share identical master‑detail smart markers, creates a DataSet with Orders and OrderDetails tables, binds them to a WorkbookDesigner, processes the markers while preserving unknown tags, and saves each populated workbook to a designated output folder.
    class Program
    {
        static void Main()
        {
            try
            {
                // List of template file paths (Excel files that contain the same master‑detail smart markers)
                var templateFiles = new List<string>
                {
                    @"Templates\ReportTemplate1.xlsx",
                    @"Templates\ReportTemplate2.xlsx",
                    @"Templates\ReportTemplate3.xlsx"
                };

                // Output folder for the generated reports
                string outputFolder = @"GeneratedReports\";

                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Prepare master‑detail data sources
                DataSet dataSet = CreateMasterDetailDataSet();

                // Process each template
                foreach (string templatePath in templateFiles)
                {
                    // Verify that the template file exists
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
                        WorkbookDesigner designer = new WorkbookDesigner
                        {
                            Workbook = workbook,
                            // When using a named range for smart markers set LineByLine to false
                            LineByLine = false
                        };

                        // Bind the master and detail tables to the designer
                        designer.SetDataSource(dataSet.Tables["Orders"]);
                        designer.SetDataSource(dataSet.Tables["OrderDetails"]);

                        // Process the smart markers (true = preserve unrecognized markers)
                        designer.Process(true);

                        // Build output file name based on the template name
                        string outputPath = Path.Combine(
                            outputFolder,
                            Path.GetFileNameWithoutExtension(templatePath) + "_Result.xlsx");

                        // Save the processed workbook
                        workbook.Save(outputPath);
                        Console.WriteLine($"Processed and saved: {outputPath}");
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
        private static DataSet CreateMasterDetailDataSet()
        {
            DataSet ds = new DataSet();

            // Master table: Orders
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("CustomerName", typeof(string));
            orders.Columns.Add("OrderDate", typeof(DateTime));
            orders.Rows.Add(1001, "Acme Corp", DateTime.Today.AddDays(-10));
            orders.Rows.Add(1002, "Beta Ltd.", DateTime.Today.AddDays(-5));
            ds.Tables.Add(orders);

            // Detail table: OrderDetails
            DataTable details = new DataTable("OrderDetails");
            details.Columns.Add("OrderID", typeof(int));
            details.Columns.Add("Product", typeof(string));
            details.Columns.Add("Quantity", typeof(int));
            details.Columns.Add("UnitPrice", typeof(decimal));
            details.Rows.Add(1001, "Widget A", 10, 9.99m);
            details.Rows.Add(1001, "Widget B", 5, 19.99m);
            details.Rows.Add(1002, "Gadget X", 2, 49.99m);
            details.Rows.Add(1002, "Gadget Y", 1, 99.99m);
            ds.Tables.Add(details);

            return ds;
        }
    }
}
