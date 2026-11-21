// Title: Create a smart marker in Aspose.Cells for .NET that fills a chart data range and enables automatic chart updates
// Common Searches: Aspose.Cells C# smart marker to populate Excel chart series from a DataTable | How to generate an auto‑updating column chart using smart markers in .NET | Using WorkbookDesigner to expand chart data range with smart markers in Aspose.Cells | Dynamic Excel chart range with smart markers and DataTable in C# | Create smart marker placeholders for chart categories and values in Aspose.Cells
// Tags: smart marker chart data range | WorkbookDesigner process smart markers | dynamic column chart Aspose.Cells | populate chart series from DataTable | auto updating Excel chart with smart markers

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SmartMarkerChartExample
{
    // The example creates a new workbook, adds a DataTable with quarterly sales figures, inserts smart markers (&=Category, &=Value) into the worksheet, defines a column chart whose series and category ranges reference those marker cells, processes the markers with WorkbookDesigner to expand the data rows, and saves the file. The resulting Excel file contains a chart that automatically updates when the underlying data changes.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet and rename it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Prepare sample data in a DataTable
                DataTable dt = new DataTable("SalesData");
                dt.Columns.Add("Category", typeof(string));
                dt.Columns.Add("Value", typeof(double));

                dt.Rows.Add("Q1", 12000);
                dt.Rows.Add("Q2", 15000);
                dt.Rows.Add("Q3", 18000);
                dt.Rows.Add("Q4", 21000);

                // Insert headers
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Insert smart markers that will be replaced with the data rows
                // Smart marker syntax: &=ColumnName
                sheet.Cells["A2"].PutValue("&=Category");
                sheet.Cells["B2"].PutValue("&=Value");

                // Add a column chart that will use the data populated by smart markers
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Define the series and category ranges – they will expand after smart markers are processed
                string seriesRange = $"Data!$B$2:$B${dt.Rows.Count + 1}";
                string categoryRange = $"Data!$A$2:$A${dt.Rows.Count + 1}";

                // Add the series and set category (X) axis data
                chart.NSeries.Add(seriesRange, true);
                chart.NSeries.CategoryData = categoryRange;

                // Optional: set chart title
                chart.Title.Text = "Quarterly Sales";

                // Process smart markers (lifecycle rule: load & process)
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource(dt);
                designer.Process();

                // Save the workbook (lifecycle rule: save)
                workbook.Save("SmartMarkerChartOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
