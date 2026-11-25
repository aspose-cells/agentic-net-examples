// Title: Hide chart legend on a dashboard worksheet to maximize plot area with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide legend in column chart on dashboard sheet | remove chart legend to enlarge plot area using Aspose.Cells .NET | set chart.Legend.IsDeleted true Aspose.Cells example | adjust plot area size after hiding legend in Aspose.Cells chart | create dashboard worksheet with column chart and no legend Aspose.Cells
// Tags: chart legend removal Aspose.Cells C# | dashboard sheet chart layout Aspose.Cells | disable automatic plot area sizing Aspose.Cells | column chart without legend Aspose.Cells | increase chart display area Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook with a "Dashboard" worksheet, adds sample month‑sales data, inserts a column chart, hides the chart legend by setting chart.Legend.IsDeleted to true, disables automatic plot‑area sizing, sets explicit width and height to use the freed space, and saves the file as DashboardChart_HideLegend.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a worksheet that will act as the dashboard
        Worksheet dashboard = workbook.Worksheets[0];
        dashboard.Name = "Dashboard";

        // Populate some sample data for the chart
        // Category labels
        dashboard.Cells["A1"].PutValue("Month");
        dashboard.Cells["A2"].PutValue("Jan");
        dashboard.Cells["A3"].PutValue("Feb");
        dashboard.Cells["A4"].PutValue("Mar");
        dashboard.Cells["A5"].PutValue("Apr");

        // Series 1 values
        dashboard.Cells["B1"].PutValue("Sales");
        dashboard.Cells["B2"].PutValue(120);
        dashboard.Cells["B3"].PutValue(150);
        dashboard.Cells["B4"].PutValue(180);
        dashboard.Cells["B5"].PutValue(200);

        // Add a column chart to the dashboard sheet
        // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
        int chartIndex = dashboard.Charts.Add(ChartType.Column, 7, 0, 25, 10);
        Chart chart = dashboard.Charts[chartIndex];

        // Set the data range for the chart
        // The first parameter is the range that includes both categories and series values
        chart.SetChartDataRange("A1:B5", true);

        // Hide the legend to maximize the display area
        // The Legend object has an IsDeleted property; setting it to true removes the legend from the chart
        chart.Legend.IsDeleted = true;

        // Optional: adjust the plot area to use the freed space (disable automatic sizing)
        chart.PlotArea.IsAutomaticSize = false;
        chart.PlotArea.Width = 8000;   // width in points (adjust as needed)
        chart.PlotArea.Height = 4000;  // height in points (adjust as needed)

        // Save the workbook to a file
        string outputPath = "DashboardChart_HideLegend.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook saved to '{outputPath}'. Legend has been hidden.");
    }
}
