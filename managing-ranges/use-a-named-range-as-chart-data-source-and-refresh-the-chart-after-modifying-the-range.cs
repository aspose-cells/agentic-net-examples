// Title: Aspose.Cells .NET – Bind a Chart to a Named Range and Refresh After Data Changes
// Description: This example demonstrates how to create a named range, assign it as the data source for a column chart, modify the underlying cells, and force the chart to refresh using Chart.Calculate. It also shows how to verify the update with IsChartDataChanged before saving the workbook.
// Keywords: Aspose.Cells | C# chart named range | Chart.Calculate | IsChartDataChanged | .NET workbook example | dynamic chart data source | refresh chart after data change | named range chart series | Aspose.Cells sample code | GitHub Aspose.Cells chart example
// Common Searches: Aspose.Cells bind chart to named range C# | refresh chart after modifying data Aspose.Cells | Chart.Calculate method usage Aspose.Cells .NET | detect chart data change Aspose.Cells | sample code named range chart Aspose.Cells
// Developer Intent: Create a chart that reads data from a named range, change the range values programmatically, and update the chart to reflect those changes.
// Use Cases: Generate a column chart whose series points to a named range, enabling easy data updates without rebuilding the chart. | Programmatically alter values inside the named range and call Chart.Calculate to automatically redraw the chart. | Validate that the chart recognized the modification by checking IsChartDataChanged before exporting the workbook.
// AI Prompts: Show how to bind a chart series to a named range and refresh it after data changes using Aspose.Cells for .NET. | Provide C# code that updates a named range, recalculates the chart, and verifies the change with IsChartDataChanged. | Explain the steps to create a named range, assign it to a chart, and ensure the chart updates when the range values are modified.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsNamedRangeChartDemo
{
    // This example demonstrates how to create a named range, assign it as the data source for a column chart, modify the underlying cells, and force the chart to refresh using Chart.Calculate. It also shows how to verify the update with IsChartDataChanged before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Create a named range that refers to the values column (B2:B4)
                int nameIndex = workbook.Worksheets.Names.Add("MyData");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                namedRange.RefersTo = "=Sheet1!$B$2:$B$4";

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Use the named range as the data source for the series
                chart.NSeries.Add("MyData", true);

                // Note: CategoryData property may not be available in some versions.
                // Categories will be taken from the first column by default.

                // Save the workbook with the initial chart
                string initialPath = "Chart_With_NamedRange_Initial.xlsx";
                workbook.Save(initialPath);
                Console.WriteLine($"Workbook saved: {Path.GetFullPath(initialPath)}");

                // ----- Modify the data behind the named range -----
                sheet.Cells["B3"].PutValue(45); // Change value for category B

                // Refresh the chart data. Calculate forces the chart to re‑read its source.
                chart.Calculate();

                // Optional: verify that the chart detects the change
                bool changed = chart.IsChartDataChanged();
                Console.WriteLine("Chart data changed flag after modification: " + changed);

                // Save the workbook after the data change and chart refresh
                string updatedPath = "Chart_With_NamedRange_Updated.xlsx";
                workbook.Save(updatedPath);
                Console.WriteLine($"Workbook saved: {Path.GetFullPath(updatedPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
