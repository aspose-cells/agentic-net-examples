// Title: Batch generate Excel reports from multiple templates with master‑detail smart markers using Aspose.Cells in C#
// AI Prompts: Write a C# console app that loads an array of Excel template files, assigns a shared master‑detail DataSet to WorkbookDesigner, processes all smart markers, and saves each workbook as a separate report, handling missing templates gracefully. | Create code that iterates over a list of template paths, creates a WorkbookDesigner for each workbook, reuses the same DataSet as the smart‑marker data source, ensures the output folder exists, and writes the processed files to disk.
// Common Searches: c# aspose.cells apply same smart markers to several Excel templates | batch generate master‑detail reports with smart markers using Aspose.Cells | iterate over multiple workbook templates and process smart markers in C# | reuse a DataSet for smart markers across many Excel files | handle missing template files when generating reports with Aspose.Cells
// Tags: batch smart marker processing Aspose.Cells | master-detail dataset smart markers C# | process multiple Excel templates Aspose.Cells | WorkbookDesigner reuse DataSet | generate reports from template workbooks C#

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace BatchSmartMarkerReport
{
    // The example builds a master‑detail DataSet, loops through an array of Excel template files (creating a blank workbook when a template is missing), assigns the DataSet to WorkbookDesigner for each workbook, processes all smart markers, ensures the output directory exists, and saves each generated report to a specified file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare a master‑detail DataSet that will be used for all templates
                DataSet reportData = CreateMasterDetailDataSet();

                // List of template files and corresponding output files
                string[] templateFiles = { "Template1.xlsx", "Template2.xlsx", "Template3.xlsx" };
                string[] outputFiles   = { "Report1.xlsx",   "Report2.xlsx",   "Report3.xlsx" };

                // Iterate over each template, apply the same smart markers and generate the report
                for (int i = 0; i < templateFiles.Length; i++)
                {
                    Workbook workbook;

                    // Load the template workbook if it exists; otherwise create a new workbook
                    if (File.Exists(templateFiles[i]))
                    {
                        workbook = new Workbook(templateFiles[i]);
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Template file '{templateFiles[i]}' not found. Creating a blank workbook.");
                        workbook = new Workbook(); // creates a default workbook with one worksheet
                    }

                    // Initialize the designer with the loaded workbook
                    WorkbookDesigner designer = new WorkbookDesigner(workbook);

                    // Set the same master‑detail data source for the current workbook
                    designer.SetDataSource(reportData);

                    // Process all smart markers in the workbook
                    designer.Process();

                    // Ensure the output directory exists
                    string outputPath = Path.GetFullPath(outputFiles[i]);
                    string outputDir = Path.GetDirectoryName(outputPath);
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Save the generated report
                    workbook.Save(outputFiles[i]);
                    Console.WriteLine($"Report generated: {outputFiles[i]}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Creates a DataSet containing a master table (Orders) and a detail table (OrderDetails)
        private static DataSet CreateMasterDetailDataSet()
        {
            DataSet ds = new DataSet();

            // Master table
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("OrderID", typeof(int));
            orders.Columns.Add("Customer", typeof(string));
            orders.Rows.Add(1, "Acme Corp");
            orders.Rows.Add(2, "Globex Inc");
            ds.Tables.Add(orders);

            // Detail table
            DataTable orderDetails = new DataTable("OrderDetails");
            orderDetails.Columns.Add("OrderID", typeof(int));
            orderDetails.Columns.Add("Product", typeof(string));
            orderDetails.Columns.Add("Quantity", typeof(int));
            orderDetails.Rows.Add(1, "Widget", 10);
            orderDetails.Rows.Add(1, "Gadget", 5);
            orderDetails.Rows.Add(2, "Doohickey", 7);
            ds.Tables.Add(orderDetails);

            // Define relation between master and detail tables
            DataRelation relation = new DataRelation(
                "Orders_OrderDetails",
                orders.Columns["OrderID"],
                orderDetails.Columns["OrderID"]);
            ds.Relations.Add(relation);

            return ds;
        }
    }
}
