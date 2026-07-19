// Title: Create a Chart Sheet with an OHLC Stock Chart using Aspose.Cells for .NET (C#)
// Description: This example shows how to generate a new workbook, fill a worksheet with Date, Open, High, Low and Close values, add a dedicated chart sheet, and insert a StockOpenHighLowClose (OHLC) chart that references the data range. The chart is titled and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | OHLC chart | StockOpenHighLowClose | chart sheet | Excel financial chart | example code | tutorial | GitHub sample | API usage
// Common Searches: Aspose.Cells add OHLC chart on separate sheet | C# create chart sheet with StockOpenHighLowClose | How to plot Open High Low Close data using Aspose.Cells | Aspose.Cells example for financial chart generation | Create chart sheet in Excel with Aspose.Cells .NET
// Developer Intent: Add a dedicated chart sheet that displays an OHLC stock chart derived from worksheet data.
// Use Cases: Generate financial reports where each security’s OHLC chart is isolated on its own sheet for clarity. | Automate batch creation of multiple chart sheets, each representing a different instrument or time period. | Export ready‑to‑share Excel workbooks containing accurate OHLC visualizations for analysts.
// AI Prompts: Provide C# code that uses Aspose.Cells to add a chart sheet and insert a StockOpenHighLowClose chart from a given data range. | Show how to create several OHLC chart sheets in one workbook, each linked to its own data table. | Explain how to customize the OHLC chart title, axis labels, and series colors with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example shows how to generate a new workbook, fill a worksheet with Date, Open, High, Low and Close values, add a dedicated chart sheet, and insert a StockOpenHighLowClose (OHLC) chart that references the data range. The chart is titled and the workbook is saved as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Use the first worksheet to hold sample OHLC data
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Header row
        dataSheet.Cells["A1"].PutValue("Date");
        dataSheet.Cells["B1"].PutValue("Open");
        dataSheet.Cells["C1"].PutValue("High");
        dataSheet.Cells["D1"].PutValue("Low");
        dataSheet.Cells["E1"].PutValue("Close");

        // Sample data rows (Day 1 to Day 5)
        for (int i = 2; i <= 6; i++)
        {
            dataSheet.Cells[$"A{i}"].PutValue($"Day {i - 1}");
            dataSheet.Cells[$"B{i}"].PutValue(10 + i); // Open
            dataSheet.Cells[$"C{i}"].PutValue(15 + i); // High
            dataSheet.Cells[$"D{i}"].PutValue(8 + i);  // Low
            dataSheet.Cells[$"E{i}"].PutValue(12 + i); // Close
        }

        // Add a new chart sheet (type Chart)
        int chartSheetIdx = workbook.Worksheets.Add(SheetType.Chart);
        Worksheet chartSheet = workbook.Worksheets[chartSheetIdx];
        chartSheet.Name = "OHLC Chart";

        // Add an OHLC (Open‑High‑Low‑Close) stock chart to the chart sheet
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
        int chartIdx = chartSheet.Charts.Add(ChartType.StockOpenHighLowClose, 5, 5, 25, 15);
        Chart chart = chartSheet.Charts[chartIdx];

        // Define the data range for the chart (Open, High, Low, Close)
        // The series are plotted by column (isVertical = true)
        chart.NSeries.Add("Data!B2:E6", true);
        // Set the category (X‑axis) data to the dates column
        chart.NSeries.CategoryData = "Data!A2:A6";

        // Optional: give the chart a title
        chart.Title.Text = "OHLC Stock Chart";

        // Save the workbook
        workbook.Save("OHLCChart.xlsx", SaveFormat.Xlsx);
    }
}
