// Title: Add High/Low Data Labels to a Stock High‑Low‑Close Chart with Aspose.Cells for .NET
// Description: Generate an Excel workbook, fill it with date‑open‑high‑low‑close data, create a StockHighLowClose chart, enable data labels for the High and Low series, position high labels above the points and low labels below, then save the file using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# stock chart | StockHighLowClose | data labels | high low values | label position above below | Excel chart automation | financial charting | .NET Excel library
// Common Searches: Aspose.Cells show high and low values on a stock chart | C# add data labels to StockHighLowClose chart | position stock chart data labels above and below in Aspose.Cells | enable data labels for multiple series in Aspose.Cells .NET | how to label high/low points in Excel stock chart using code
// Developer Intent: Enable and position data labels for the High and Low series of a StockHighLowClose chart.
// Use Cases: Automated generation of financial reports where daily high and low prices are displayed directly on the chart for quick visual reference. | Creating investor‑ready Excel workbooks that highlight price extremes on stock charts without manual editing. | Building a market‑data export tool that programmatically adds readable high/low labels to compliance‑oriented Excel charts.
// AI Prompts: Write C# code with Aspose.Cells to add data labels to a StockHighLowClose chart, show high and low values, and set high labels above and low labels below the points. | Provide an Aspose.Cells example that customizes the font color and size of high and low data labels in a stock chart. | Explain how to programmatically enable data labels for multiple series and control their positions in an Excel stock chart using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Generate an Excel workbook, fill it with date‑open‑high‑low‑close data, create a StockHighLowClose chart, enable data labels for the High and Low series, position high labels above the points and low labels below, then save the file using Aspose.Cells in C#.
class StockChartDataLabels
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data: Date, Open, High, Low, Close
            dataSheet.Cells["A1"].PutValue("Date");
            dataSheet.Cells["B1"].PutValue("Open");
            dataSheet.Cells["C1"].PutValue("High");
            dataSheet.Cells["D1"].PutValue("Low");
            dataSheet.Cells["E1"].PutValue("Close");

            dataSheet.Cells["A2"].PutValue("01-Jan");
            dataSheet.Cells["B2"].PutValue(100);
            dataSheet.Cells["C2"].PutValue(110);
            dataSheet.Cells["D2"].PutValue(95);
            dataSheet.Cells["E2"].PutValue(105);

            dataSheet.Cells["A3"].PutValue("02-Jan");
            dataSheet.Cells["B3"].PutValue(105);
            dataSheet.Cells["C3"].PutValue(115);
            dataSheet.Cells["D3"].PutValue(100);
            dataSheet.Cells["E3"].PutValue(110);

            dataSheet.Cells["A4"].PutValue("03-Jan");
            dataSheet.Cells["B4"].PutValue(108);
            dataSheet.Cells["C4"].PutValue(118);
            dataSheet.Cells["D4"].PutValue(102);
            dataSheet.Cells["E4"].PutValue(112);

            // Add a chart sheet and create a Stock High‑Low‑Close chart
            Worksheet chartSheet = workbook.Worksheets[workbook.Worksheets.Add(SheetType.Chart)];
            Chart stockChart = chartSheet.Charts[chartSheet.Charts.Add(ChartType.StockHighLowClose, 5, 0, 25, 15)];

            // Add series for High and Low values
            Series highSeries = stockChart.NSeries[stockChart.NSeries.Add("C2:C4", true)];
            Series lowSeries = stockChart.NSeries[stockChart.NSeries.Add("D2:D4", true)];

            // Set the category (X‑axis) data (dates)
            stockChart.NSeries.CategoryData = "A2:A4";

            // Enable data labels and configure them to show the values (high and low)
            highSeries.DataLabels.ShowValue = true;
            lowSeries.DataLabels.ShowValue = true;

            // Position the high values above the points and low values below
            highSeries.DataLabels.Position = LabelPositionType.Above;
            lowSeries.DataLabels.Position = LabelPositionType.Below;

            // Save the workbook
            workbook.Save("StockChartWithHighLowDataLabels.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
