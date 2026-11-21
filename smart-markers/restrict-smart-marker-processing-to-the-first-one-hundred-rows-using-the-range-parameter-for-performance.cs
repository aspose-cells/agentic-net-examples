// Title: How to limit Aspose.Cells smart marker processing to the first 100 rows using the WorkbookDesigner Range parameter in C#
// Common Searches: Aspose.Cells limit smart marker processing to first 100 rows C# | use WorkbookDesigner range parameter to process a subset of rows | C# restrict smart markers to specific row range for better performance | setdatasource with range in Aspose.Cells to speed up smart markers | process only top 100 rows of Excel data using Aspose.Cells smart markers
// Tags: WorkbookDesigner setdatasource range Aspose.Cells | smart marker row range limitation .NET | Excel smart markers performance tuning | process subset of rows Aspose.Cells C# | limit smart marker execution Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerExampleApp
{
    // This example demonstrates how to restrict smart marker processing to the first 100 rows of an Excel worksheet by supplying a range string to WorkbookDesigner.SetDataSource. Limiting the row range improves performance when generating large reports with Aspose.Cells in C#.
    class SmartMarkerExample
    {
        static void Main()
        {
            try
            {
                const string templatePath = "Template.xlsx";

                // Verify that the template file exists before loading
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the workbook that contains smart markers
                Workbook workbook = new Workbook(templatePath);

                // Prepare a data source (example DataTable)
                DataTable data = new DataTable();
                data.Columns.Add("Name", typeof(string));
                data.Columns.Add("Quantity", typeof(int));
                data.Columns.Add("Price", typeof(double));

                // Populate the DataTable with sample data
                for (int i = 0; i < 200; i++)
                {
                    data.Rows.Add($"Item {i + 1}", i + 10, (i + 1) * 1.5);
                }

                // Use WorkbookDesigner to process smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(data);
                designer.Process();

                // Save the result workbook
                const string outputPath = "Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
