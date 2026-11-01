// Title: Generate multiple Excel reports by processing master‑detail smart markers across several template workbooks with Aspose.Cells for .NET
// Common Searches: c# batch generate Excel reports using Aspose.Cells smart markers from a DataSet | process master detail smart markers in multiple template files with Aspose.Cells WorkbookDesigner | how to loop through Excel templates and apply the same smart markers in C# | Aspose.Cells generate reports for each template in a folder using smart markers | automate Excel report creation for many templates with master‑detail data source in .NET
// Tags: batch smart marker processing Aspose.Cells | WorkbookDesigner master-detail data binding | C# create Excel workbooks from DataSet | iterate template workbooks Aspose.Cells | automated Excel report generation .NET

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace BatchReportGenerator
{
    // Iterates over a collection of Excel template files, loads each workbook, uses WorkbookDesigner to bind a master‑detail DataSet as the smart‑marker data source, processes the smart markers, and saves a generated report for each template.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // List of template file paths
                List<string> templateFiles = new List<string>
                {
                    @"C:\Templates\ReportTemplate1.xlsx",
                    @"C:\Templates\ReportTemplate2.xlsx",
                    // Add more template paths as needed
                };

                // Output directory
                string outputDir = @"C:\Reports\";

                // Ensure the output directory exists
                Directory.CreateDirectory(outputDir);

                // Prepare master-detail data source
                DataSet reportData = CreateMasterDetailDataSet();

                // Process each template
                foreach (string templatePath in templateFiles)
                {
                    // Verify template file exists
                    if (!File.Exists(templatePath))
                    {
                        Console.WriteLine($"Template not found: {templatePath}");
                        continue;
                    }

                    try
                    {
                        // Load the template workbook
                        Workbook workbook = new Workbook(templatePath);

                        // Create a WorkbookDesigner to work with smart markers
                        WorkbookDesigner designer = new WorkbookDesigner(workbook);

                        // Set the data source for the smart markers (master-detail)
                        designer.SetDataSource(reportData);

                        // Process the smart markers
                        designer.Process();

                        // Determine output file name
                        string fileName = Path.GetFileNameWithoutExtension(templatePath);
                        string outputPath = Path.Combine(outputDir, $"{fileName}_Generated.xlsx");

                        // Save the generated report
                        workbook.Save(outputPath);

                        Console.WriteLine($"Generated report: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing template '{templatePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch reports generation completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Creates a DataSet containing master and detail tables for the smart markers
        private static DataSet CreateMasterDetailDataSet()
        {
            DataSet ds = new DataSet();

            // Master table (e.g., Orders)
            DataTable masterTable = new DataTable("Orders");
            masterTable.Columns.Add("OrderID", typeof(int));
            masterTable.Columns.Add("CustomerName", typeof(string));
            masterTable.Columns.Add("OrderDate", typeof(DateTime));

            // Detail table (e.g., OrderDetails)
            DataTable detailTable = new DataTable("OrderDetails");
            detailTable.Columns.Add("OrderID", typeof(int)); // Foreign key to master
            detailTable.Columns.Add("Product", typeof(string));
            detailTable.Columns.Add("Quantity", typeof(int));
            detailTable.Columns.Add("UnitPrice", typeof(decimal));

            // Populate master rows
            masterTable.Rows.Add(1001, "Alice Johnson", new DateTime(2023, 5, 12));
            masterTable.Rows.Add(1002, "Bob Smith", new DateTime(2023, 5, 13));
            masterTable.Rows.Add(1003, "Carol Lee", new DateTime(2023, 5, 14));

            // Populate detail rows (multiple rows per master)
            detailTable.Rows.Add(1001, "Laptop", 1, 1200.00m);
            detailTable.Rows.Add(1001, "Mouse", 2, 25.50m);
            detailTable.Rows.Add(1002, "Desk", 1, 300.00m);
            detailTable.Rows.Add(1002, "Chair", 4, 85.00m);
            detailTable.Rows.Add(1003, "Monitor", 2, 200.00m);
            detailTable.Rows.Add(1003, "Keyboard", 1, 45.00m);
            detailTable.Rows.Add(1003, "Headset", 1, 60.00m);

            // Add tables to the DataSet
            ds.Tables.Add(masterTable);
            ds.Tables.Add(detailTable);

            // Define the relation between master and detail (optional but useful for smart markers)
            DataRelation relation = new DataRelation(
                "Orders_Details",
                masterTable.Columns["OrderID"],
                detailTable.Columns["OrderID"]);
            ds.Relations.Add(relation);

            return ds;
        }
    }
}
