// Title: Batch processing multiple Excel templates with smart markers using Aspose.Cells in C#
// Common Searches: c# batch generate Excel reports from a list of template files using Aspose.Cells WorkbookDesigner | load several Excel templates and apply smart markers with different DataTables in Aspose.Cells | process multiple workbook templates programmatically and save populated files to a folder | use Aspose.Cells to bind DataTable collections to templates for monthly report generation
// Tags: batch workbook processing with WorkbookDesigner | smart markers population for multiple Excel templates | Aspose.Cells generate reports from DataTable collection | save populated workbooks as Xlsx programmatically | C# load template workbook and bind data source

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace BatchWorkbookProcessor
{
    // The example iterates over a list of Excel template paths, creates matching DataTable objects, loads each template into a Workbook, binds the DataTable to a WorkbookDesigner to process smart markers, and saves the populated workbook with a month‑based filename to a specified output directory.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define template files and corresponding data sources
                var templates = new List<string>
                {
                    @"C:\Templates\ReportTemplate1.xlsx",
                    @"C:\Templates\ReportTemplate2.xlsx",
                    @"C:\Templates\ReportTemplate3.xlsx"
                };

                // Example data sources – in real scenarios these could come from a database, service, etc.
                var dataSources = new List<DataTable>
                {
                    CreateSampleDataTable("January"),
                    CreateSampleDataTable("February"),
                    CreateSampleDataTable("March")
                };

                // Output folder for populated workbooks
                string outputFolder = @"C:\GeneratedReports";

                // Ensure the output folder exists
                Directory.CreateDirectory(outputFolder);

                // Process each template with its matching data source
                for (int i = 0; i < templates.Count; i++)
                {
                    string templatePath = templates[i];
                    DataTable data = dataSources[i];

                    if (!File.Exists(templatePath))
                    {
                        Console.WriteLine($"Template not found: {templatePath}");
                        continue; // Skip to next iteration
                    }

                    try
                    {
                        // Load the workbook template
                        Workbook workbook = new Workbook(templatePath);

                        // Use WorkbookDesigner to bind the data source and process smart markers
                        WorkbookDesigner designer = new WorkbookDesigner(workbook);
                        designer.SetDataSource(data);
                        designer.Process(); // Populates the workbook with data

                        // Build the output file name (e.g., Report_January.xlsx)
                        string monthName = data.Rows[0]["Month"].ToString();
                        string outputPath = Path.Combine(outputFolder, $"Report_{monthName}.xlsx");

                        // Save the populated workbook
                        workbook.Save(outputPath, SaveFormat.Xlsx);

                        Console.WriteLine($"Generated report: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing template '{templatePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Helper method to create a sample DataTable for demonstration purposes
        private static DataTable CreateSampleDataTable(string month)
        {
            DataTable table = new DataTable("SalesData");
            table.Columns.Add("Month", typeof(string));
            table.Columns.Add("Product", typeof(string));
            table.Columns.Add("Quantity", typeof(int));
            table.Columns.Add("Revenue", typeof(double));

            // Add sample rows
            table.Rows.Add(month, "Product A", 120, 1500.00);
            table.Rows.Add(month, "Product B", 85, 1120.50);
            table.Rows.Add(month, "Product C", 60, 780.25);

            return table;
        }
    }
}
