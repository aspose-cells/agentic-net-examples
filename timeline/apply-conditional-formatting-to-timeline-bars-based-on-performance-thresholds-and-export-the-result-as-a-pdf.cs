// Title: Color timeline column chart bars by performance thresholds and export as PDF using Aspose.Cells for .NET
// Common Searches: how to change column chart bar colors based on cell values with Aspose.Cells C# | apply conditional formatting to individual chart points in Aspose.Cells .NET | export a timeline chart with colored bars to PDF using Aspose.Cells | set bar fill color according to performance value in Aspose.Cells column chart | create a performance‑based timeline report PDF in C# with Aspose.Cells
// Tags: chart series point color conditional formatting Aspose.Cells | timeline column chart performance thresholds C# | export workbook to PDF with chart Aspose.Cells | set bar fill color based on cell value Aspose.Cells | conditional bar coloring in Aspose.Cells chart

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, populates dates and performance values, adds a column chart that acts as a timeline, colors each bar green, yellow, or red based on performance thresholds (≥80, ≥50, otherwise), and saves the worksheet with the chart as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample timeline data: dates in column A, performance values in column B
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(startDate.AddDays(i));   // Date
                cells[i, 1].PutValue(30 + i * 10);            // Performance value (30,40,...,120)
            }

            // Add a column chart (acts as a timeline bar chart)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 12, 0, 26, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B1:B10", true);          // Values
            chart.NSeries.CategoryData = "A1:A10";     // Categories (dates)

            // Apply conditional formatting to each data point based on performance thresholds
            Series series = chart.NSeries[0];
            for (int i = 0; i < series.Points.Count; i++)
            {
                try
                {
                    // Retrieve the corresponding cell value (performance) for the point
                    double value = cells[i, 1].DoubleValue;
                    Color fillColor;

                    if (value >= 80)
                        fillColor = Color.Green;   // High performance
                    else if (value >= 50)
                        fillColor = Color.Yellow;  // Medium performance
                    else
                        fillColor = Color.Red;     // Low performance

                    // Set the fill color for the individual bar (data point)
                    series.Points[i].Area.ForegroundColor = fillColor;
                }
                catch (Exception pointEx)
                {
                    Console.WriteLine($"Point formatting error at index {i}: {pointEx.Message}");
                }
            }

            // Export the worksheet (including the chart) to PDF
            workbook.Save("TimelineReport.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
