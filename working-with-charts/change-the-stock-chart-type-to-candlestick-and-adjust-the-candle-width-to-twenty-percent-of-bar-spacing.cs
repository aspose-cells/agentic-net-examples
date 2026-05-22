using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data required for a stock chart (Date, Open, High, Low, Close)
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Open");
        sheet.Cells["C1"].PutValue("High");
        sheet.Cells["D1"].PutValue("Low");
        sheet.Cells["E1"].PutValue("Close");

        sheet.Cells["A2"].PutValue("2023-01-01");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["C2"].PutValue(110);
        sheet.Cells["D2"].PutValue(95);
        sheet.Cells["E2"].PutValue(105);

        sheet.Cells["A3"].PutValue("2023-01-02");
        sheet.Cells["B3"].PutValue(105);
        sheet.Cells["C3"].PutValue(115);
        sheet.Cells["D3"].PutValue(100);
        sheet.Cells["E3"].PutValue(110);

        sheet.Cells["A4"].PutValue("2023-01-03");
        sheet.Cells["B4"].PutValue(110);
        sheet.Cells["C4"].PutValue(120);
        sheet.Cells["D4"].PutValue(108);
        sheet.Cells["E4"].PutValue(115);

        // Add a stock chart (initially High‑Low‑Close) to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.StockHighLowClose, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the series (Open, High, Low, Close) and categories (Date)
        chart.NSeries.Add("B2:E4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Change the chart type to candlestick (Open‑High‑Low‑Close)
        chart.Type = ChartType.StockOpenHighLowClose;

        // Adjust candle width to 20 % of the bar spacing using the GapWidth property
        chart.GapWidth = 20; // 20 % gap between candles

        // Save the workbook with the modified chart
        workbook.Save("CandlestickChart.xlsx");
    }
}