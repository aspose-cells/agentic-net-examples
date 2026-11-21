// Title: Create a grouped sales report with region subtotals using Aspose.Cells smart markers in C#
// Common Searches: how to use Aspose.Cells WorkbookDesigner to group rows by region and add subtotal rows in C# | C# smart markers example for aggregating sales data and inserting subtotals in Excel | Aspose.Cells generate Excel report with grouped data and subtotals using smart markers | group sales data by region with subtotal rows using Aspose.Cells smart markers C# | process DataSet with smart markers to create grouped Excel report in .NET
// Tags: WorkbookDesigner grouping smart markers | smart markers subtotal rows C# | region based sales aggregation Aspose.Cells | DataSet smart marker processing .NET | Excel report generation with grouped data Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerExample
{
    // The example loads a template workbook, builds a DataTable with Region and Sales columns, adds it to a DataSet, assigns the DataSet to a WorkbookDesigner, processes smart markers that contain grouping and subtotal directives, and saves the resulting Excel file with sales data grouped by region and subtotal rows.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the template workbook
                const string templatePath = "Template.xlsx";

                // Verify that the template file exists to avoid FileNotFoundException
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Error: Template file '{templatePath}' was not found.");
                    return;
                }

                // Load the workbook that contains the smart marker template
                Workbook workbook = new Workbook(templatePath);

                // Prepare the sales data
                DataTable salesTable = new DataTable("Sales");
                salesTable.Columns.Add("Region", typeof(string));
                salesTable.Columns.Add("Sales", typeof(double));

                // Sample data – in a real scenario this could come from a database or other source
                salesTable.Rows.Add("North", 1200.00);
                salesTable.Rows.Add("North", 800.00);
                salesTable.Rows.Add("South", 1500.00);
                salesTable.Rows.Add("South", 700.00);
                salesTable.Rows.Add("East", 900.00);
                salesTable.Rows.Add("West", 1100.00);

                // Put the DataTable into a DataSet (required by the smart marker processor)
                DataSet dataSet = new DataSet();
                dataSet.Tables.Add(salesTable);

                // Process the smart markers using WorkbookDesigner (compatible with all Aspose.Cells versions)
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource(dataSet);
                designer.Process();

                // Save the resulting workbook with grouped data and subtotal rows
                const string resultPath = "Result.xlsx";

                // Ensure the directory for the result file exists
                string resultDir = Path.GetDirectoryName(Path.GetFullPath(resultPath));
                if (!Directory.Exists(resultDir))
                {
                    Directory.CreateDirectory(resultDir);
                }

                workbook.Save(resultPath);
                Console.WriteLine($"Workbook processed successfully and saved as '{resultPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
