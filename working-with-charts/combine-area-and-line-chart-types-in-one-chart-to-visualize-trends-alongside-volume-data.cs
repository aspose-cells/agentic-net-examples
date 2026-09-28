// Title: Generate an Excel combo chart that displays volume as an Area series and trend as a Line series using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, fills it with month, volume, and trend data, adds an Area chart, and then changes the second series to a Line chart to form a combo visualization. | Demonstrate how to programmatically set different chart types for individual series in an Aspose.Cells chart and save the result as an XLSX file.
// Common Searches: aspocells c# create combo chart with area and line series | how to change a series to line type in Aspose.Cells chart | example of mixed area and line chart using Aspose.Cells .NET | Aspose.Cells set chart series type programmatically C# | generate Excel chart showing sales volume as area and forecast as line with Aspose.Cells
// Tags: Aspose.Cells create combo area line chart | Aspose.Cells set series type programmatically | C# Aspose.Cells mixed chart types | Aspose.Cells generate Excel chart with area and line series | Aspose.Cells chart series type conversion

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program builds a new workbook, inserts month, volume, and trend data, adds an Area chart covering the data range, converts the second series to a Line chart to produce a combo chart, sets a title and legend, and saves the workbook as ComboAreaLineChart.xlsx.
class ComboAreaLineChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data:
        // Column A – Category (e.g., months)
        // Column B – Volume (to be shown as Area)
        // Column C – Trend (to be shown as Line)
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Volume");
        sheet.Cells["C1"].PutValue("Trend");

        string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        double[] volume = { 120, 150, 180, 130, 170, 200 };
        double[] trend = { 100, 130, 160, 110, 150, 190 };

        for (int i = 0; i < months.Length; i++)
        {
            sheet.Cells[i + 1, 0].PutValue(months[i]);   // A column
            sheet.Cells[i + 1, 1].PutValue(volume[i]);  // B column
            sheet.Cells[i + 1, 2].PutValue(trend[i]);   // C column
        }

        // Add a chart. Start with Area type; we'll later change the second series to Line.
        int chartIndex = sheet.Charts.Add(ChartType.Area, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart (including all three columns)
        chart.SetChartDataRange("A1:C7", true);

        // By default, all series inherit the chart type (Area). Change the second series (Trend) to Line.
        // Series collection order follows the order of data columns (B then C).
        // Series[0] corresponds to Volume (Area), Series[1] corresponds to Trend (Line).
        chart.NSeries[1].Type = ChartType.Line;

        // Optional: give the chart a title and enable legend
        chart.Title.Text = "Volume (Area) and Trend (Line) Combo Chart";
        chart.ShowLegend = true;

        // Save the workbook to an XLSX file
        workbook.Save("ComboAreaLineChart.xlsx");
    }
}
