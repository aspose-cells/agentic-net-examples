// Title: Aspose.Cells for .NET – Add and customize data labels on a Stock Open‑High‑Low‑Close chart
// Description: C# example that creates a workbook, fills it with OHLC data, inserts a StockOpenHighLowClose chart, enables data labels, positions them above each point, and replaces the default label with a custom text showing the high and low values (e.g., "H:120 L:80"). The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# stock chart | Open‑High‑Low‑Close chart | custom data labels | show high low values | data label position above points | Excel chart labeling | .NET chart series formatting | financial chart automation
// Common Searches: Aspose.Cells display high low values on stock chart | C# add data labels to StockOpenHighLowClose chart | customize stock chart labels Aspose.Cells .NET | set data label position above points Aspose.Cells | replace default label with custom text in Aspose.Cells chart
// Developer Intent: Create a stock chart and configure its series to show a custom label that combines the high and low values for each data point.
// Use Cases: Generate financial reports with OHLC data where each point on a stock chart displays "H:{high} L:{low}" for quick visual analysis. | Update an existing Excel workbook programmatically to improve readability of stock charts by adding concise high/low labels. | Automate Excel chart creation for trading dashboards, positioning labels above points to avoid overlap with the chart grid.
// AI Prompts: Write C# code using Aspose.Cells to insert a StockOpenHighLowClose chart and set data labels that show "H:{high} L:{low}" for each point. | Show how to loop through series points in an Aspose.Cells stock chart and assign a custom label while hiding the default value. | Explain how to change the data label position to above the points for a stock chart series in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, fills it with OHLC data, inserts a StockOpenHighLowClose chart, enables data labels, positions them above each point, and replaces the default label with a custom text showing the high and low values (e.g., "H:120 L:80"). The workbook is saved as an XLSX file.
    public class StockChartDataLabelsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a stock chart
            // Column A: Date (category)
            // Column B: High values
            // Column C: Low values
            // Column D: Open values
            // Column E: Close values
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("High");
            sheet.Cells["C1"].PutValue("Low");
            sheet.Cells["D1"].PutValue("Open");
            sheet.Cells["E1"].PutValue("Close");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["C2"].PutValue(80);
            sheet.Cells["D2"].PutValue(100);
            sheet.Cells["E2"].PutValue(110);

            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(130);
            sheet.Cells["C3"].PutValue(85);
            sheet.Cells["D3"].PutValue(115);
            sheet.Cells["E3"].PutValue(120);

            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(125);
            sheet.Cells["C4"].PutValue(90);
            sheet.Cells["D4"].PutValue(110);
            sheet.Cells["E4"].PutValue(115);

            // Add a stock chart (Open‑High‑Low‑Close)
            int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose, 6, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the stock series (Open, High, Low, Close)
            chart.NSeries.Add("B2:E4", true);
            // Set category (X‑axis) data – the dates
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first (and only) series
            Series series = chart.NSeries[0];

            // Enable data labels for the series
            series.DataLabels.ShowValue = true; // Show the numeric values
            series.DataLabels.Position = LabelPositionType.Above; // Position labels above the points

            // Customize each point's label to show High and Low values
            for (int i = 0; i < series.Points.Count; i++)
            {
                // Retrieve High and Low values for the current point
                double high = sheet.Cells[i + 1, 1].DoubleValue; // B column (High)
                double low = sheet.Cells[i + 1, 2].DoubleValue;  // C column (Low)

                // Compose custom label text
                series.Points[i].DataLabels.Text = $"H:{high} L:{low}";
                series.Points[i].DataLabels.ShowValue = false; // Hide default value, use custom text
            }

            // Save the workbook
            workbook.Save("StockChartDataLabelsDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
