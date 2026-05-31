using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load (or create) a workbook
        Workbook workbook = new Workbook(); // creates a new workbook
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a scatter chart
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells["A" + i].PutValue(i);                         // X values
            sheet.Cells["B" + i].PutValue(Math.Pow(i, 2) + 5);        // Y values (quadratic with offset)
        }

        // Add a scatter chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Define the series data
        chart.NSeries.Add("B2:B6", true);          // Y values
        chart.NSeries[0].XValues = "A2:A6";       // X values
        chart.NSeries[0].Name = "Sample Series";

        // Add a polynomial trendline of order 3 to the first series
        int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Polynomial);
        Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];
        trendline.Order = 3;               // Set polynomial order to 3
        trendline.DisplayEquation = true;  // Show the equation on the chart
        trendline.DisplayRSquared = true;  // Optionally show R‑squared

        // Save the workbook
        string filePath = "PolynomialTrendline.xlsx";
        workbook.Save(filePath);

        // Reload the workbook to demonstrate capturing trendline information
        Workbook loadedWorkbook = new Workbook(filePath);
        Chart loadedChart = loadedWorkbook.Worksheets[0].Charts[0];
        Trendline loadedTrendline = loadedChart.NSeries[0].TrendLines[0];

        // Capture and output relevant properties
        Console.WriteLine("Trendline Order: " + loadedTrendline.Order);
        Console.WriteLine("Display Equation: " + loadedTrendline.DisplayEquation);
        // Note: Aspose.Cells does not expose the actual equation string via API.
    }
}