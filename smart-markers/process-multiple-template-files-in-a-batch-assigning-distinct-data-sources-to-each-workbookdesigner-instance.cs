// Title: Batch generate Excel files from multiple smart‑marker templates with individual data sources using Aspose.Cells WorkbookDesigner (C#)
// AI Prompts: Write C# code that loads a list of Excel template paths, creates a WorkbookDesigner for each, sets a unique data source (e.g., DataTable) with SetDataSource, processes the smart markers, and saves the result to a specified folder. | Refactor the batch processor to use a Dictionary<string, object> that maps each template file name to its corresponding data object, then iterate over the dictionary to apply each data source to its WorkbookDesigner. | Add robust error handling and logging to the batch loop so that missing templates or processing exceptions are recorded and the loop continues with the next file.
// Common Searches: how to use Aspose.Cells WorkbookDesigner to process several Excel templates in one C# program | batch create Excel reports from different DataTables with smart markers using Aspose.Cells | assign different data sources to multiple WorkbookDesigner instances in C# | loop through Excel template files and generate output files with Aspose.Cells smart markers | C# example for processing a collection of Excel templates with separate data sources
// Tags: batch processing Excel templates with WorkbookDesigner | Aspose.Cells assign distinct data source per template | smart markers multiple workbook generation C# | WorkbookDesigner SetDataSource for each file | generate Excel files from DataTable collection Aspose.Cells

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace BatchTemplateProcessor
{
    // Iterates over a collection of Excel template files, creates a WorkbookDesigner for each, assigns a unique data source, processes smart markers, and saves the generated workbooks to an output directory.
    class Program
    {
        static void Main(string[] args)
        {
            // Define template files and corresponding data sources
            var templates = new List<string>
            {
                @"C:\Templates\InvoiceTemplate.xlsx",
                @"C:\Templates\ReportTemplate.xlsx",
                @"C:\Templates\SummaryTemplate.xlsx"
            };

            // Example data sources: each entry can be any object (DataTable, List<T>, etc.)
            var dataSources = new List<object>
            {
                GetInvoiceData(),
                GetReportData(),
                GetSummaryData()
            };

            // Output folder
            string outputFolder = @"C:\GeneratedFiles\";

            // Ensure the output directory exists
            try
            {
                Directory.CreateDirectory(outputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create output directory '{outputFolder}': {ex.Message}");
                return;
            }

            // Process each template with its data source
            for (int i = 0; i < templates.Count; i++)
            {
                string templatePath = templates[i];
                object dataSource = dataSources[i];

                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: '{templatePath}'. Skipping.");
                    continue;
                }

                try
                {
                    // Load the template workbook
                    Workbook workbook = new Workbook(templatePath);

                    // Create a WorkbookDesigner for the loaded workbook
                    WorkbookDesigner designer = new WorkbookDesigner(workbook);

                    // Assign the data source to the designer.
                    // The name "DataSource" must match the name used in the template markers.
                    designer.SetDataSource("DataSource", dataSource);

                    // Process the template (merge data with markers)
                    designer.Process();

                    // Build output file name
                    string outputPath = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(templatePath) + "_Result.xlsx");

                    // Save the processed workbook
                    workbook.Save(outputPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Processed '{templatePath}' and saved to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing template '{templatePath}': {ex.Message}");
                }
            }
        }

        // Mock methods to provide sample data sources.
        // Replace these with actual data retrieval logic.

        private static DataTable GetInvoiceData()
        {
            DataTable dt = new DataTable("Invoice");
            dt.Columns.Add("Item", typeof(string));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("Price", typeof(decimal));

            dt.Rows.Add("Widget A", 10, 2.5m);
            dt.Rows.Add("Widget B", 5, 5.0m);
            dt.Rows.Add("Widget C", 2, 12.0m);

            return dt;
        }

        private static DataTable GetReportData()
        {
            DataTable dt = new DataTable("Report");
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns.Add("Metric", typeof(string));
            dt.Columns.Add("Value", typeof(double));

            dt.Rows.Add(DateTime.Today.AddDays(-2), "Sales", 1500.75);
            dt.Rows.Add(DateTime.Today.AddDays(-1), "Sales", 1720.30);
            dt.Rows.Add(DateTime.Today, "Sales", 1890.10);

            return dt;
        }

        private static DataTable GetSummaryData()
        {
            DataTable dt = new DataTable("Summary");
            dt.Columns.Add("Category", typeof(string));
            dt.Columns.Add("Total", typeof(decimal));

            dt.Rows.Add("North Region", 12500.00m);
            dt.Rows.Add("South Region", 9800.50m);
            dt.Rows.Add("East Region", 11200.75m);
            dt.Rows.Add("West Region", 13450.20m);

            return dt;
        }
    }
}
