// Title: Aspose.Cells C# – Export Chart to PNG with Data Labels Visible
// Description: Creates a workbook, fills it with sample data, adds a column chart, enables value labels, and uses the ToImage method to save the chart as a PNG file while keeping the data labels displayed in the image.
// Keywords: Aspose.Cells | C# chart export | PNG image from chart | preserve data labels | ToImage method | column chart | Excel to image conversion | chart snapshot | Aspose.Cells API
// Common Searches: Aspose.Cells export chart as PNG with labels | C# keep data labels when saving chart image | ToImage chart data labels visible | How to render chart with values in PNG using Aspose.Cells | Save Excel chart to PNG preserving labels
// Developer Intent: Generate a PNG image of an Excel chart that retains its data‑label values.
// Use Cases: Embedding chart snapshots with numeric values into web dashboards. | Automating report generation where charts must show exact figures. | Creating email or PDF assets that require chart images with visible labels.
// AI Prompts: Write C# code with Aspose.Cells to export a chart to PNG while showing data labels. | Explain the steps to enable chart data labels before calling ToImage in Aspose.Cells. | Provide an example that exports multiple charts to separate PNG files, each preserving its labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExport
{
    // Creates a workbook, fills it with sample data, adds a column chart, enables value labels, and uses the ToImage method to save the chart as a PNG file while keeping the data labels displayed in the image.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(1500);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Enable data labels and make sure values are shown
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Convert the chart to PNG while preserving data labels
            chart.ToImage("ChartWithLabels.png", ImageType.Png);

            // Optionally save the workbook (not required for image export)
            workbook.Save("ChartWorkbook.xlsx");
        }
    }
}
