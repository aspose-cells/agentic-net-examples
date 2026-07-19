// Title: Auto‑Refresh Excel Chart on Workbook Open with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a column chart, and use OoxmlSaveOptions.RefreshChartCache so Excel automatically updates the chart each time the XLSX file is opened.
// Keywords: Aspose.Cells chart auto refresh | RefreshChartCache .NET | OoxmlSaveOptions chart update | C# Excel chart refresh on open | Aspose.Cells automatic chart update
// Common Searches: Aspose.Cells enable chart refresh when opening file | C# set RefreshChartCache for Excel chart | auto update chart data Aspose.Cells | save workbook with chart auto‑refresh option
// Developer Intent: Configure a workbook so its embedded chart recalculates automatically when the file is opened in Excel.
// Use Cases: Daily sales dashboard that always shows the latest figures without manual refresh. | Financial model workbooks where charts must reflect updated calculations on each open. | Automated reporting pipelines that generate Excel files and need charts to display current data instantly.
// AI Prompts: How do I enable automatic chart refresh on workbook open using Aspose.Cells for .NET? | Provide C# code that creates a chart and sets RefreshChartCache so the chart updates when the XLSX is opened. | Explain the limitations of RefreshChartCache in Aspose.Cells and how it interacts with Excel's calculation engine.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAutoRefreshDemo
{
    // Demonstrates how to create a workbook, add a column chart, and use OoxmlSaveOptions.RefreshChartCache so Excel automatically updates the chart each time the XLSX file is opened.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable automatic refresh of the chart cache when the workbook is opened
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
                {
                    RefreshChartCache = true // Instruct Excel to refresh chart data on open
                };

                // Save the workbook with the specified options
                string outputPath = "ChartAutoRefreshOnOpen.xlsx";
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook saved to '{outputPath}'. Chart will refresh automatically when opened in Excel.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
